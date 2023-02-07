using Odey.Framework.Keeley.Entities.Enums;
using Odey.Framework.Keeley.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using E = Odey.Framework.Keeley.Entities;
using Odey.Excel.CrispinsSpreadsheet.Data_Access;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class FXExposureManager
    {
        private readonly List<E.Portfolio> _portfolio;


        private readonly Fund _fund;


        public FXExposureManager(List<E.Portfolio> portfolio, Fund fund)
        {
            _portfolio = portfolio;
            _fund = fund;
        }

        private List<FXExposure> Build()
        {
            return _portfolio.GroupBy(g => new
            {
                Currency = g.Position.Currency,
                BookId = g.Position.BookID,
                InstrumentMarket = g.Position.Currency.Instrument.InstrumentMarkets.First(),
                IsAccrual = false,
                ReferenceDate = g.ReferenceDate,
                Label = GetLabel(g.Position.InstrumentMarket, out var fXExposureTypeId),
                FXExposureType = fXExposureTypeId
            })
                .Select(s => new FXExposure(
                    s.Key.Currency,
                    s.Key.BookId,
                    s.Key.InstrumentMarket,
                    s.Key.Label,
                    s.Key.FXExposureType,
                    s.Key.ReferenceDate,
                    s.Sum(a => a.MarketValue / a.FXRate),
                    s.Average(a => a.FXRate),
                    s.Sum(a => a.MarketValue)
                    )).ToList();

        }
        private List<FXExposure> _fxExposure = null;
        private List<FXExposure> FXExposure
        {
            get
            {
                if (_fxExposure == null)
                {
                    _fxExposure = Build();
                }
                return _fxExposure;
            }

        }

        private string GetLabel(InstrumentMarket instrumentMarket, out FXExposureTypeIds fxExposureTypeId)
        {
            switch (instrumentMarket.InstrumentClassIdAsEnum)
            {
                case InstrumentClassIds.ForwardFX:
                    if (instrumentMarket.Instrument.DerivedAssetClassId == (int)DerivedAssetClassIds.ForeignExchange)
                    {
                        fxExposureTypeId = FXExposureTypeIds.HedgeFX;
                        var label = instrumentMarket.Instrument.Name.Substring(0, 12);
                        return label;
                    }
                    else
                    {
                        fxExposureTypeId = FXExposureTypeIds.HedgeFX;
                        return "Hedges";
                    }
                case InstrumentClassIds.InflationLinkedBond:
                    fxExposureTypeId = FXExposureTypeIds.InflationLinkedBonds;
                    return "Linkers";
                case InstrumentClassIds.GovtBond:
                    if (instrumentMarket.Name.StartsWith("UKT"))
                    {
                        fxExposureTypeId = FXExposureTypeIds.Gilts;
                        return "Gilts";
                    }
                    fxExposureTypeId = FXExposureTypeIds.OtherGovernmentBonds;
                    return "Over Govt Bonds";
                case InstrumentClassIds.Miscellaneous:
                    fxExposureTypeId = FXExposureTypeIds.Miscellaneous;
                    return "Misc";
                case InstrumentClassIds.InterestRateSwap:
                    fxExposureTypeId = FXExposureTypeIds.IRS;
                    return "IRS";
                case InstrumentClassIds.ContractForDifference:
                    if (instrumentMarket.Name.StartsWith("UKT"))
                    {
                        fxExposureTypeId = FXExposureTypeIds.GiltSwaps;
                        return "Gilts (Swap)";
                    }
                    break;

            }
            switch ((DerivedAssetClassIds) instrumentMarket.Instrument.DerivedAssetClassId)
            {
                case DerivedAssetClassIds.Cash:
                    fxExposureTypeId = FXExposureTypeIds.Cash;
                    return "Cash & Equivalents";
                default:
                    fxExposureTypeId = FXExposureTypeIds.Equity;
                    return "Equity";

            }
            
        }

        public List<E.Portfolio> GetUnhedged()
        {
            var fxExposure = FXExposure;
            var navs = _portfolio.GroupBy(a=>a.ReferenceDate).ToDictionary(a=>a.Key,a=>a.Sum(s => s.MarketValue));
            var toReturn = fxExposure.Where(a => a.CurrencyId != _fund.CurrencyId && a.FXExposureTypeId != FXExposureTypeIds.PropFX)
                .GroupBy(g => new
                {
                    Currency = g.Currency,
                    BookId = g.BookId,
                    InstrumentMarket = g.InstrumentMarket,                    
                    ReferenceDate = g.ReferenceDate
                })
                .Where(w => Math.Abs(w.Sum(a => a.MarketValue)) / navs[w.Key.ReferenceDate] > .02m)
                    .Select(s => new Portfolio()
                    {
                        Position = new E.Position()
                        {
                            CurrencyID = s.Key.Currency.InstrumentID,
                            Currency = s.Key.Currency,
                            BookID = s.Key.BookId,
                            InstrumentMarketID = s.Key.InstrumentMarket.InstrumentMarketID,
                            InstrumentMarket = s.Key.InstrumentMarket,
                            IsAccrual = false,
                        },
                        ReferenceDate = s.Key.ReferenceDate,
                        NetPosition = s.Sum(a => a.NetPosition),
                        Price = s.Average(a => a.Price)
                    }).ToList();
            return toReturn;
        } 

        public List<E.Portfolio> GetUnhedgedOld(DateTime referenceDate)
        {
            var nav = _portfolio.Where(a => a.ReferenceDate == referenceDate).Sum(a => a.MarketValue);
            var toReturn = _portfolio.Where(a => a.Position.CurrencyID != _fund.CurrencyId && a.Position.InstrumentMarket.Instrument.DerivedAssetClassId != (int)DerivedAssetClassIds.ForeignExchange)
                    .GroupBy(g => new
                    {
                        CurrencyID = g.Position.CurrencyID,
                        Currency = g.Position.Currency,
                        BookID = g.Position.BookID,
                        Book = g.Position.Book,
                        InstrumentMarketID = g.Position.Currency.Instrument.InstrumentMarkets.First().InstrumentMarketID,
                        InstrumentMarket = g.Position.Currency.Instrument.InstrumentMarkets.First(),
                        IsAccrual = false,
                        ReferenceDate = g.ReferenceDate
                    })
                    .Where(w => Math.Abs(w.Sum(a => a.MarketValue)) / nav > .02m)
                    .Select(s => new Portfolio()
                    {
                        Position = new E.Position()
                        {
                            CurrencyID = s.Key.CurrencyID,
                            Currency = s.Key.Currency,
                            BookID = s.Key.BookID,
                            Book = s.Key.Book,
                            InstrumentMarketID = s.Key.InstrumentMarketID,
                            InstrumentMarket = s.Key.InstrumentMarket,
                            IsAccrual = false,
                        },
                        ReferenceDate = s.Key.ReferenceDate,
                        NetPosition = s.Sum(a => a.MarketValue / a.FXRate),
                        Price = s.Average(a => a.FXRate)
                    }).ToList();


            return toReturn;
        }
    }
}
