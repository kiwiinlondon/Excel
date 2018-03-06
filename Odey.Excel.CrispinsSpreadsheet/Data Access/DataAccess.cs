using Odey.Framework.Keeley.Entities;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class DataAccess
    {
        public DataAccess(DateTime referenceDate)
        {
            ReferenceDate = referenceDate;
            PreviousReferenceDate = GetPreviousReferenceDate(referenceDate);
            PreviousPreviousReferenceDate = GetPreviousReferenceDate(PreviousReferenceDate);
        }

        public DateTime PreviousPreviousReferenceDate { get; private set; }
        public DateTime PreviousReferenceDate { get; private set; }
        public DateTime ReferenceDate { get; private set; }

       

        private decimal? GetPrice(IEnumerable<Portfolio> positions, InstrumentTypeIds tickerTypeId)
        {
            positions = positions.Where(a => a != null);
            if (positions == null || positions.Count()==0)
            {
                return null;
            }

            var prices = positions.Select(a => a.Price * a.Position.InstrumentMarket.PriceQuoteMultiplier).Distinct();

            if (prices.Count()!=1 && tickerTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                throw new ApplicationException("Cannot establish unique price");
            }
            return prices.First();
        }

        


        private DateTime GetPreviousReferenceDate(DateTime referenceDate)
        {
            DateTime previousReferenceDate = referenceDate.AddDays(-1);
            if (previousReferenceDate.DayOfWeek == DayOfWeek.Sunday)
            {
                return previousReferenceDate.AddDays(-2);
            }
            return previousReferenceDate;
        }

        
        public List<PortfolioDTO> GetPortfolio(Fund fund, bool includeHedging, bool onlyIncludeFX)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                DateTime[] referenceDates = { PreviousPreviousReferenceDate, PreviousReferenceDate, ReferenceDate };


                var portfolios = context.Portfolios
                    .Include(a => a.Position.Book)
                    .Include(a => a.Position.Currency.Instrument)
                    .Include(a => a.Position.InstrumentMarket.PriceCurrency.Instrument)
                    .Include(a => a.Position.InstrumentMarket.Instrument)
                    .Include(a => a.Position.InstrumentMarket.Market.LegalEntity.Country)
                    .Where(a => a.FundId == fund.FundId && referenceDates.Contains(a.ReferenceDate) && a.Position.IsAccrual == false && !a.IsFlat);

                if (!includeHedging)//Share 
                {
                     portfolios = portfolios.Where(a=>a.Position.InstrumentMarket.Instrument.DerivedAssetClassId != (int)DerivedAssetClassIds.Cash);
                }
                if (onlyIncludeFX)//MAC and share classes
                {
                    portfolios = portfolios.Where(a => a.Position.InstrumentMarket.Instrument.InstrumentClassID == (int)InstrumentClassIds.ForwardFX);
                }

                var portfoliosByPosition = portfolios.ToList()
                    .GroupBy(g => g.Position)
                    .Select(a => new DTOGrouping
                    {
                        Position = a.Key,
                        PreviousPrevious = a.FirstOrDefault(f => f.ReferenceDate == PreviousPreviousReferenceDate),
                        Previous = a.FirstOrDefault(f => f.ReferenceDate == PreviousReferenceDate),
                        Current = a.FirstOrDefault(f => f.ReferenceDate == ReferenceDate),
                    });

                var fxPositions = portfoliosByPosition.Where(a => a.Position.InstrumentMarket.InstrumentClassIdAsEnum == InstrumentClassIds.ForwardFX).ToList();
                var fxToAdd = BuildFX(fxPositions, fund);
                return portfoliosByPosition
                    .Where(a => a.Position.InstrumentMarket.InstrumentClassIdAsEnum != InstrumentClassIds.ForwardFX)
                    .GroupBy(g => new { Book = g.Position.Book.Name, Instrument = InstrumentBuilder.Instance.Get(g.Position.InstrumentMarket) })
                .Select(a => new PortfolioDTO(
                    a.Key.Book,
                    a.Key.Instrument,
                    a.Sum(s => s.Previous == null ? 0 : s.Previous.NetPosition),
                    a.Sum(s=> s.Current == null ? 0 : s.Current.NetPosition),                   
                    GetPrice(a.Select(s => s.PreviousPrevious), a.Key.Instrument.InstrumentTypeId),
                    GetPrice(a.Select(s => s.Previous), a.Key.Instrument.InstrumentTypeId),
                    GetPrice(a.Select(s=>s.Current), a.Key.Instrument.InstrumentTypeId)
                    ))
                    .Union(fxToAdd).ToList();
            }
        }

       

        private List<PortfolioDTO> BuildFX(List<DTOGrouping> fxPositions, Fund fund)
        {
            var quantities = fxPositions
                    .GroupBy(g => new { Book = g.Position.Book.Name, Instrument = InstrumentBuilder.Instance.GetFX(g.Position.InstrumentMarket) })
                    .Select(a =>
                    new
                    {
                        Key = a.Key,
                        PreviousPreviousNetPosition = GetFXNetPosition(a.Select(s => s.PreviousPrevious),a.Key.Instrument),
                        PreviousNetPosition = GetFXNetPosition(a.Select(s => s.Previous), a.Key.Instrument),
                        CurrentNetPosition = GetFXNetPosition(a.Select(s => s.Current), a.Key.Instrument)
                    }
                    );
            return quantities.Select(a => new PortfolioDTO(a.Key.Book, a.Key.Instrument,a.PreviousNetPosition, a.CurrentNetPosition, null, null, null)
            ).ToList();
        }

        private decimal GetFXNetPosition(IEnumerable<Portfolio> positions, InstrumentDTO instrument)
        {
            positions = positions.Where(a => a != null && a.NetPosition !=0);
            if (positions == null || positions.Count() == 0)
            {
                return 0;
            }

            var nonFlat = positions.GroupBy(a => new { a.Position.BookID, a.Position.InstrumentMarketID, a.Position.AccountID, a.Position.Strategy })
                .Where(a => a.Count() >= 2).SelectMany(a => a);
            return nonFlat.Where(a => a.Position.Currency.IsoCode == instrument.Currency).Sum(a => a.NetPosition);            
        }

       

        public Fund GetFund(FundIds fundId)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var fund = context.Funds.Include(a=>a.LegalEntity).Include(a=>a.Currency.Instrument).FirstOrDefault(a => a.LegalEntityID == (int)fundId);
                Fund toReturn = new Fund(fund.LegalEntityID,fund.Name,fund.Currency.IsoCode,false);
                DateTime[] referenceDates = { PreviousReferenceDate, ReferenceDate };
                var navs = context.FundNetAssetValues.Where(a => a.FundId == fund.LegalEntityID && referenceDates.Contains(a.ReferenceDate)).ToList();
                toReturn.Nav = navs.FirstOrDefault(a => a.ReferenceDate == ReferenceDate).MarketValue;
                toReturn.PreviousNav = navs.FirstOrDefault(a => a.ReferenceDate == PreviousReferenceDate).MarketValue;                
                return toReturn;
            }
        }



        public List<Book> GetBooks(Fund fund)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var books = context.Books.Where(a => a.FundID == fund.FundId);
                var booksById = books.ToDictionary(a => a.BookID, a => new Book(fund,a.BookID, a.Name,false));
                DateTime[] referenceDates = { PreviousReferenceDate, ReferenceDate };
                var navs = context.BookNetAssetValues.Where(a => booksById.Keys.Contains(a.BookId) && referenceDates.Contains(a.ReferenceDate)).ToList();
                foreach (var nav in navs)
                {
                    var book = booksById[nav.BookId];
                    if (nav.ReferenceDate == PreviousReferenceDate)
                    {
                        book.PreviousNav = nav.MarketValue.Value;
                    }
                    else if (nav.ReferenceDate == ReferenceDate)
                    {
                        book.Nav = nav.MarketValue.Value;
                    }
                    else
                    {
                        throw new ApplicationException($"Unknown date {nav.ReferenceDate}");
                    }

                }
                return booksById.Values.ToList();
            }
        }




        public List<FXRateDTO> GetFXRates()
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                DateTime[] referenceDates = { PreviousPreviousReferenceDate, PreviousReferenceDate };
                var todayRates = context.FXRates.Include(a => a.FromCurrency.Instrument).Include(a => a.ToCurrency.Instrument)
                    .Where(a => referenceDates.Contains(a.ReferenceDate) && a.ReferenceDate == a.ForwardDate);
                return todayRates.GroupBy(g => new { FromCurrency = g.FromCurrency.Instrument.Name, ToCurrency = g.ToCurrency.Instrument.Name })
                      .Select(a => new FXRateDTO()
                      {
                          FromCurrency = a.Key.FromCurrency,
                          ToCurrency = a.Key.ToCurrency,
                          PreviousPreviousValue = a.FirstOrDefault(f => f.ReferenceDate == PreviousPreviousReferenceDate).Value,
                          PreviousValue = a.FirstOrDefault(f => f.ReferenceDate == PreviousReferenceDate).Value
                      }).ToList();
            }
        }

        public InstrumentDTO GetInstrument(string ticker)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var instrumentMarkets = context.InstrumentMarkets
                    .Include(a=>a.Instrument)
                    .Include(a => a.Market.LegalEntity.Country)
                .Where(a => a.BloombergTicker == ticker).ToList();
                if (instrumentMarkets.Count == 0)
                {
                    return null;
                }
                if (instrumentMarkets.Count >1)
                {
                    instrumentMarkets = instrumentMarkets.Where(a => a.InstrumentClassIdAsEnum != InstrumentClassIds.ContractForDifference).ToList();
                }
                return InstrumentBuilder.Instance.Get(instrumentMarkets.FirstOrDefault());
            }
        }


        public void AddExchangeCountryToInstrument(InstrumentDTO instrument)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                string ticker = instrument.Identifier.Code;
                int countOfSpaces = instrument.Identifier.Code.Count(a => a == ' ');
                if (countOfSpaces==2)
                {
                    string exchangeCode = ticker.Substring(ticker.IndexOf(' '));
                    exchangeCode = exchangeCode.Substring(0, exchangeCode.IndexOf(' ') - 1);
                    var market = context.Markets.Include(a=>a.Country).FirstOrDefault(a => a.BBExchangeCode == exchangeCode);
                    if (market!=null)
                    {
                        instrument.ExchangeCountryIsoCode = market.Country.IsoCode;
                        instrument.ExchangeCountryName = market.Country.Name;
                    }
                }
            }
        }
    }
}
