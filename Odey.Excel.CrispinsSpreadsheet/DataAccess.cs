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

        private decimal GetPrice(IEnumerable<Portfolio> positions)
        {
            var prices = positions.Select(a => a.Price * a.Position.InstrumentMarket.PriceQuoteMultiplier).Distinct();
            if (prices.Count()!=1)
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

        public List<PortfolioDTO> Get(int fundId, DateTime referenceDate, out decimal nav)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                nav = context.FundNetAssetValues.FirstOrDefault(a => a.FundId == fundId && a.ReferenceDate == referenceDate).MarketValue;

                var positions = context.Portfolios
                    .Include(a => a.Position.InstrumentMarket.PriceCurrency.Instrument)
                    .Include(a => a.Position.InstrumentMarket.Instrument)
                    .Include(a=>a.Position.InstrumentMarket.Market.LegalEntity.Country)
                    .Where(a => a.FundId == fundId && a.ReferenceDate == referenceDate && a.Position.IsAccrual == false 
                        && !AssetClassIdsToExclude.Contains(a.Position.InstrumentMarket.Instrument.DerivedAssetClassId) && !a.IsFlat).ToList();
                return positions.GroupBy(g => new
                {
                    CountryIso = g.Position.InstrumentMarket.Market.LegalEntity.Country.IsoCode,
                    CountryName = g.Position.InstrumentMarket.Market.LegalEntity.Country.Name,
                    Name = g.Position.InstrumentMarket.Name,
                    Ticker = GetTicker(g.Position.InstrumentMarket),
                    Currency = g.Position.InstrumentMarket.PriceCurrency.IsoCode,
                    PriceDivisor = g.Position.InstrumentMarket.PriceDivisor
                })
                .Select(a => new PortfolioDTO(a.Key.Name,a.Key.Ticker.Item1,a.Key.Currency, a.Key.CountryIso,a.Key.CountryName,a.Sum(s=>s.NetPosition), a.Key.Ticker.Item2, GetPrice(a),a.Key.PriceDivisor)).ToList();
            }
        }

    }
}
