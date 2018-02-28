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

        public static readonly int[] AssetClassIdsToExclude = new int[] { (int)DerivedAssetClassIds.Cash };

        private decimal? GetPrice(IEnumerable<Portfolio> positions, TickerTypeIds tickerTypeId)
        {
            positions = positions.Where(a => a != null);
            if (positions == null || positions.Count()==0)
            {
                return null;
            }

            var prices = positions.Select(a => a.Price * a.Position.InstrumentMarket.PriceQuoteMultiplier).Distinct();

            if (prices.Count()!=1 && tickerTypeId == TickerTypeIds.PrivatePlacement)
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

        private string GetName(IEnumerable<DTOGrouping> groupings)
        {

            var names = groupings.Select(a => a.Position.InstrumentMarket.Name).Distinct();
            if (names.Count() > 1)
            {
                names = groupings.Where(a => a.Position.InstrumentMarket.InstrumentClassIdAsEnum != InstrumentClassIds.ContractForDifference)
                    .Select(a => a.Position.InstrumentMarket.Name).Distinct();

                if (names.Count() > 1)
                {
                    throw new ApplicationException($"Cannot establish unique name. Names = {string.Join(",", names)}");
                }
            }
            return names.First();
        }

        public List<PortfolioDTO> Get(FundDTO fund)
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

                if (fund.FundFXTreatmentId != FundFXTreatmentIds.ShareClass)//MAC & OEI
                {
                     portfolios = portfolios.Where(a=>!AssetClassIdsToExclude.Contains(a.Position.InstrumentMarket.Instrument.DerivedAssetClassId));
                }
                if (fund.FundFXTreatmentId != FundFXTreatmentIds.Normal)//MAC and share classes
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
                    .GroupBy(g => DTOGroupBuilder.Instance.Get(g.Position, fund))
                .Select(a => new PortfolioDTO(
                    a.Key.Book,
                    a.Key.AssetClass,
                    GetName(a),
                    a.Key.Ticker,
                    a.Key.Currency, 
                    a.Key.CountryIso,
                    a.Key.CountryName,
                    a.Sum(s => s.Previous == null ? 0 : s.Previous.NetPosition),
                    a.Sum(s=> s.Current == null ? 0 : s.Current.NetPosition), 
                    a.Key.TickerTypeId,
                    GetPrice(a.Select(s => s.PreviousPrevious), a.Key.TickerTypeId),
                    GetPrice(a.Select(s => s.Previous), a.Key.TickerTypeId),
                    GetPrice(a.Select(s=>s.Current), a.Key.TickerTypeId),
                    a.Key.PriceDivisor))
                    .Union(fxToAdd).ToList();
            }
        }

        private string GetFXName(string ticker)
        {
            return $"{ticker.Substring(0, 3)}/{ticker.Substring(3, 3)}";
        }

        private List<PortfolioDTO> BuildFX(List<DTOGrouping> fxPositions, FundDTO fund)
        {
            var quantities = fxPositions
                    .GroupBy(g => DTOGroupBuilder.Instance.GetFX(g.Position, fund))
                    .Select(a =>
                    new
                    {
                        Key = a.Key,
                        PreviousPreviousNetPosition = GetFXNetPosition(a.Select(s => s.PreviousPrevious),a.Key),
                        PreviousNetPosition = GetFXNetPosition(a.Select(s => s.Previous), a.Key),
                        CurrentNetPosition = GetFXNetPosition(a.Select(s => s.Current), a.Key)
                    }
                    );
            return quantities.Select(a => new PortfolioDTO(
                a.Key.Book, a.Key.AssetClass, GetFXName(a.Key.Ticker), a.Key.Ticker, a.Key.Currency, a.Key.CountryIso, a.Key.CountryName,
                a.PreviousNetPosition, a.CurrentNetPosition, a.Key.TickerTypeId, null, null, null, a.Key.PriceDivisor)
            ).ToList();
        }

        private decimal GetFXNetPosition(IEnumerable<Portfolio> positions, DTOGroup group)
        {
            positions = positions.Where(a => a != null && a.NetPosition !=0);
            if (positions == null || positions.Count() == 0)
            {
                return 0;
            }

            var nonFlat = positions.GroupBy(a => new { a.Position.BookID, a.Position.InstrumentMarketID, a.Position.AccountID, a.Position.Strategy })
                .Where(a => a.Count() >= 2).SelectMany(a => a);
            return nonFlat.Where(a => a.Position.Currency.IsoCode == group.Currency).Sum(a => a.NetPosition);            
        }

        private FundFXTreatmentIds GetFundFXTreatmentId(Odey.Framework.Keeley.Entities.Fund fund)
        {
            if (fund.FundTypeId == (int)FundTypeIds.ShareClass)
            {
                return FundFXTreatmentIds.ShareClass;
            }
            else if (fund.LegalEntityID != (int) FundIds.OEI)
            {
                return FundFXTreatmentIds.FXOnly;
            }
            else
            {
                return FundFXTreatmentIds.Normal;
            }
        }

        public FundDTO GetFund(FundIds fundId)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var fund = context.Funds.Include(a=>a.LegalEntity).Include(a=>a.Currency.Instrument).FirstOrDefault(a => a.LegalEntityID == (int)fundId);
                FundDTO toReturn = new FundDTO() { FundId = fund.LegalEntityID, Name = fund.Name, Currency = fund.Currency.IsoCode, FundFXTreatmentId = GetFundFXTreatmentId(fund) };
                DateTime[] referenceDates = { PreviousReferenceDate, ReferenceDate };
                var navs = context.FundNetAssetValues.Where(a => a.FundId == fund.LegalEntityID && referenceDates.Contains(a.ReferenceDate)).ToList();
                toReturn.CurrentNav = navs.FirstOrDefault(a => a.ReferenceDate == ReferenceDate).MarketValue;
                toReturn.PreviousNav = navs.FirstOrDefault(a => a.ReferenceDate == PreviousReferenceDate).MarketValue;
                if (toReturn.FundFXTreatmentId == FundFXTreatmentIds.Normal)
                {
                    AddBooks(toReturn, context, referenceDates);
                }
                return toReturn;
            }
        }

        public void AddBooks(FundDTO fund, KeeleyModel context, DateTime[] referenceDates)
        {
            var books = context.Books.Where(a => a.FundID == fund.FundId);
            var booksById = books.ToDictionary(a=>a.BookID,a => new BookDTO() { BookId = a.BookID, Name = a.Name });
           
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

            fund.Books = booksById.Values.ToDictionary(a=>a.Name,a=>a);
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
    }
}
