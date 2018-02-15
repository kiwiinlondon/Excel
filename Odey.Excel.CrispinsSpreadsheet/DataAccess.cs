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

        public static readonly int[] AssetClassIdsToInclude = new int[] { (int)DerivedAssetClassIds.Equity };

        public List<PortfolioDTO> Get(int fundId, DateTime referenceDate, out decimal nav)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                nav = context.FundNetAssetValues.FirstOrDefault(a => a.FundId == fundId && a.ReferenceDate == referenceDate).MarketValue;

                var positions = context.Portfolios.Include(a=>a.Position.InstrumentMarket.Instrument.Issuer.LegalEntity.Country).Where(a => a.FundId == fundId && a.ReferenceDate == referenceDate && a.Position.IsAccrual == false && AssetClassIdsToInclude.Contains(a.Position.InstrumentMarket.Instrument.DerivedAssetClassId)).ToList();
                return positions.GroupBy(g => new
                {
                    CountryIso = g.Position.InstrumentMarket.Instrument.Issuer.LegalEntity.Country.IsoCode,
                    CountryName = g.Position.InstrumentMarket.Instrument.Issuer.LegalEntity.Country.Name,
                    Name = g.Position.InstrumentMarket.Name,
                    Ticker = g.Position.InstrumentMarket.BloombergTicker,
                })
                .Select(a => new PortfolioDTO()
                {
                    CountryIso = a.Key.CountryIso,
                    CountryName = a.Key.CountryName,
                    Name = a.Key.Name,
                    Ticker = a.Key.Ticker,
                    NetPosition = a.Sum(s => s.NetPosition)
                }).ToList();
            }
        }

    }
}
