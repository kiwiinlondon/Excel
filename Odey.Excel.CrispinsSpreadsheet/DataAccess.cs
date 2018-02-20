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

        public static readonly int[] AssetClassIdsToExclude = new int[] { (int)DerivedAssetClassIds.ForeignExchange,(int)DerivedAssetClassIds.Cash };

        private decimal? GetPrice(IEnumerable<Portfolio> positions, int? tickerType)
        {
            positions = positions.Where(a => a != null);
            if (positions == null || positions.Count()==0)
            {
                return null;
            }

            var prices = positions.Select(a => a.Price * a.Position.InstrumentMarket.PriceQuoteMultiplier).Distinct();

            if (prices.Count()!=1 && tickerType.HasValue)
            {
                throw new ApplicationException("Cannot establish unique price");
            }
            return prices.First();
        }

        private Tuple<string,int?> GetTicker(InstrumentMarket instrumentMarket)
        {
            int? tickerType = GetTickerType(instrumentMarket);
            string ticker = instrumentMarket.BloombergTicker;
            if (tickerType.HasValue)
            {
                ticker = instrumentMarket.InstrumentMarketID.ToString();
            }
            return new Tuple<string, int?>(ticker, tickerType);
        }

        private static readonly int[] PrivateListingStatusIds = { (int)ListingStatusIds.Delisted, (int)ListingStatusIds.PrivatePlacement };

        private int? GetTickerType(InstrumentMarket instrumentMarket)
        {
            if (string.IsNullOrWhiteSpace(instrumentMarket.BloombergTicker) ||instrumentMarket.BloombergTicker.StartsWith(".") || PrivateListingStatusIds.Contains(instrumentMarket.ListingStatusId))
            {
                return 1;
            }
            return null;
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

        public List<PortfolioDTO> Get(int fundId, DateTime referenceDate, out decimal previousNav, out decimal nav)
        {
            DateTime previousReferenceDate = GetPreviousReferenceDate(referenceDate);
            DateTime previousPreviousReferenceDate= GetPreviousReferenceDate(previousReferenceDate);

            using (KeeleyModel context = new KeeleyModel())
            {
                DateTime[] referenceDates = { previousPreviousReferenceDate, previousReferenceDate, referenceDate };
                var navs = context.FundNetAssetValues.Where(a => a.FundId == fundId && referenceDates.Contains(a.ReferenceDate)).ToList();
                nav = navs.FirstOrDefault(a => a.ReferenceDate == referenceDate).MarketValue;
                previousNav = navs.FirstOrDefault(a => a.ReferenceDate == previousReferenceDate).MarketValue;

                var portfolios = context.Portfolios
                    .Include(a => a.Position.InstrumentMarket.PriceCurrency.Instrument)
                    .Include(a => a.Position.InstrumentMarket.Instrument)
                    .Include(a=>a.Position.InstrumentMarket.Market.LegalEntity.Country)
                    .Where(a => a.FundId == fundId && referenceDates.Contains( a.ReferenceDate) && a.Position.IsAccrual == false 
                        && !AssetClassIdsToExclude.Contains(a.Position.InstrumentMarket.Instrument.DerivedAssetClassId) && !a.IsFlat).ToList();

                var portfoliosByPosition = portfolios.GroupBy(g => g.Position)
                    .Select(a => new
                    {
                        Position = a.Key,
                        PreviousPrevious = a.FirstOrDefault(f => f.ReferenceDate == previousPreviousReferenceDate),
                        Previous = a.FirstOrDefault(f => f.ReferenceDate == previousReferenceDate),
                        Current = a.FirstOrDefault(f => f.ReferenceDate == referenceDate),
                    });

                return portfoliosByPosition.GroupBy(g => new
                {
                    CountryIso = g.Position.InstrumentMarket.Market.LegalEntity.Country.IsoCode,
                    CountryName = g.Position.InstrumentMarket.Market.LegalEntity.Country.Name,
                    Name = g.Position.InstrumentMarket.Name,
                    Ticker = GetTicker(g.Position.InstrumentMarket),
                    Currency = g.Position.InstrumentMarket.PriceCurrency.IsoCode,
                    PriceDivisor = g.Position.InstrumentMarket.PriceDivisor
                })
                .Select(a => new PortfolioDTO(
                    a.Key.Name,
                    a.Key.Ticker.Item1,
                    a.Key.Currency, 
                    a.Key.CountryIso,
                    a.Key.CountryName,
                    a.Sum(s => s.Previous == null ? 0 : s.Previous.NetPosition),
                    a.Sum(s=> s.Current == null ? 0 : s.Current.NetPosition), 
                    a.Key.Ticker.Item2,
                    GetPrice(a.Select(s => s.PreviousPrevious), a.Key.Ticker.Item2),
                    GetPrice(a.Select(s => s.Previous), a.Key.Ticker.Item2),
                    GetPrice(a.Select(s=>s.Current), a.Key.Ticker.Item2),
                    a.Key.PriceDivisor)).ToList();
            }
        }

    }
}
