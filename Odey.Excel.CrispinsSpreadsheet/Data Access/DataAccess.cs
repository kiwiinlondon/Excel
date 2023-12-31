﻿using Odey.Framework.Keeley.Entities;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;
using E = Odey.Framework.Keeley.Entities;

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


        private decimal? GetIndexRatio(IEnumerable<Portfolio> positions, bool isInflationAdjusted)
        {
            positions = positions.Where(a => a != null);
            if (positions == null || positions.Count() == 0)
            {
                return null;
            }

            var indexRatios = positions.Select(a => a.IndexRatio).Distinct();

            if (indexRatios.Count() != 1)
            {
                throw new ApplicationException("Cannot establish unique inflation adjustement");
            }

            var indexRatio = indexRatios.First();
           

            if (!isInflationAdjusted && indexRatio.HasValue)
            {
                throw new ApplicationException("Not expecting inflation adjustment on non inflation adjusted instrument");
            }
            return indexRatio;


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

        
        public List<PortfolioDTO> GetPortfolio(Fund fund)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                DateTime[] referenceDates = { PreviousPreviousReferenceDate, PreviousReferenceDate, ReferenceDate };


                var portfolios = context.Portfolios
                    .Include(a => a.Position.Currency.Instrument.InstrumentMarkets)
                    .Include(a => a.Position.InstrumentMarket.PriceCurrency.Instrument)
                    .Include(a => a.Position.InstrumentMarket.Instrument.InstrumentClass.ParentInstrumentClassRelationships)
                    .Include(a => a.Position.InstrumentMarket.Market.LegalEntity.Country)
                    .Include(a => a.PriceEntity.RawPrice)
                    .Where(a => a.FundId == fund.FundId
                        && referenceDates.Contains(a.ReferenceDate) && a.Position.IsAccrual == false && !a.IsFlat).ToList();

                fund.FXExposureManager = new FXExposureManager(portfolios, fund,ReferenceDate);
                var hedging = fund.FXExposureManager.GetUnhedged();
                var hedgingOld = fund.FXExposureManager.GetUnhedgedOld(ReferenceDate);
                var i = hedging.Where(a => a.Position.InstrumentMarketID == 18331).ToList();
                if (!fund.IncludeHedging)//Share 
                {
                     portfolios = portfolios.Where(a=>a.Position.InstrumentMarket.Instrument.DerivedAssetClassId != (int)DerivedAssetClassIds.Cash).ToList();
                }
                if (fund.IncludeOnlyFX)//MAC and share classes
                {
                    portfolios = portfolios.Where(a => a.Position.InstrumentMarket.Instrument.InstrumentClassID == (int)InstrumentClassIds.ForwardFX).ToList();
                }

                if (fund.Name == "OEI" || fund.Name == "SWAN")
                {
                    portfolios.AddRange(hedging);
                }

                var portfoliosByPosition = portfolios
                    .GroupBy(g => g.Position)
                    .Select(a => new DTOGrouping
                    {
                        Position = a.Key,
                        PreviousPrevious = a.FirstOrDefault(f => f.ReferenceDate == PreviousPreviousReferenceDate),
                        Previous = a.FirstOrDefault(f => f.ReferenceDate == PreviousReferenceDate),
                        Current = a.FirstOrDefault(f => f.ReferenceDate == ReferenceDate),
                    });

          
                var fxPositions = portfoliosByPosition.Where(a => a.Position.InstrumentMarket.InstrumentClassIdAsEnum == InstrumentClassIds.ForwardFX || a.Position.InstrumentMarket.InstrumentClassIdAsEnum == InstrumentClassIds.Currency).ToList();
                
                var fxToAdd = BuildFX(fxPositions, fund);
                var toReturn =  portfoliosByPosition
                    .Where(a => a.Position.InstrumentMarket.InstrumentClassIdAsEnum != InstrumentClassIds.ForwardFX && a.Position.InstrumentMarket.InstrumentClassIdAsEnum != InstrumentClassIds.Currency)
                    .GroupBy(g => new { Instrument = InstrumentBuilder.Instance.Get(g.Position.InstrumentMarket) })
                .Select(a => new PortfolioDTO(
                    a.Key.Instrument,
                    a.Sum(s => s.Previous == null ? 0 : s.Previous.NetPosition),
                    a.Sum(s=> s.Current == null ? 0 : s.Current.NetPosition),                   
                    GetPrice(a.Select(s => s.PreviousPrevious), a.Key.Instrument.InstrumentTypeId),
                    GetPrice(a.Select(s => s.Previous), a.Key.Instrument.InstrumentTypeId),
                    GetPrice(a.Select(s=>s.Current), a.Key.Instrument.InstrumentTypeId),
                    GetPriceIsManual(a.Select(s => s.PreviousPrevious)),
                    GetPriceIsManual(a.Select(s => s.Previous)),
                    GetPriceIsManual(a.Select(s => s.Current)),
                    GetIndexRatio(a.Select(s => s.PreviousPrevious),a.Key.Instrument.IsInflationAdjusted),
                    GetIndexRatio(a.Select(s => s.Previous), a.Key.Instrument.IsInflationAdjusted),
                    GetIndexRatio(a.Select(s => s.Current), a.Key.Instrument.IsInflationAdjusted)
                    ))
                    .Union(fxToAdd);
           
                return toReturn.Where(a => a.CurrentNetPosition != 0 || a.PreviousNetPosition != 0).ToList();
            }
        }



        private bool GetPriceIsManual(IEnumerable<Portfolio> positions)
        {
            positions = positions.Where(a => a != null && a.PriceEntity !=null);
            if (positions == null || positions.Count() == 0)
            {
                return false;
            }

            var prices = positions.Select(a => a.PriceEntity.RawPrice.EntityRankingSchemeItemId).Distinct().ToArray();

            var containsManual = prices.Where(a => a == (int)EntityRankingSchemeItemIds.ManualPrice || a== (int)EntityRankingSchemeItemIds.AdministratorPrice).Count();

            if (containsManual>0)
            {
                var arg = positions.FirstOrDefault(a => a.Position.InstrumentMarket.Instrument.Name.StartsWith("ARG"));
                if (arg != null)
                {
                    int i = 0;
                }
                if (prices.Length != containsManual)
                {
                    throw new ApplicationException("Cant establish whether price is manual");
                }
                return true;
            }

            return false;
        }






        private List<PortfolioDTO> BuildFX(List<DTOGrouping> fxPositions, Fund fund)
        {
            var quantities = fxPositions
                    .GroupBy(g => new {  Instrument = InstrumentBuilder.Instance.GetFX(g.Position.InstrumentMarket, fund) })
                    .Select(a =>
                    new
                    {
                        Key = a.Key,
                        PreviousPreviousNetPosition = GetFXNetPosition(a.Select(s => s.PreviousPrevious),a.Key.Instrument),
                        PreviousNetPosition = GetFXNetPosition(a.Select(s => s.Previous), a.Key.Instrument),
                        CurrentNetPosition = GetFXNetPosition(a.Select(s => s.Current), a.Key.Instrument)
                    }
                    );
            return quantities.Select(a => new PortfolioDTO(a.Key.Instrument,a.PreviousNetPosition, a.CurrentNetPosition, null, null, null, false, false, false,null,null,null)
            ).ToList();
        }

        private decimal GetFXNetPosition(IEnumerable<Portfolio> positions, InstrumentDTO instrument)
        {
            if (instrument.Name.EndsWith("ARS"))
            {
                int i = 0;
            }

            positions = positions.Where(a => a != null && a.NetPosition !=0);
            if (positions == null || positions.Count() == 0)
            {
                return 0;
            }
            if (instrument.AssetClass != "FX" )
            {
                return positions.Sum(a => a.NetPosition);
            }

            var nonFlat = positions.GroupBy(a => new { a.Position.BookID, a.Position.InstrumentMarketID, a.Position.AccountID, a.Position.Strategy })
                .Where(a => a.Count() >= 2).SelectMany(a => a);
            var netPosition = nonFlat.Where(a => a.Position.Currency.IsoCode == instrument.Currency).Sum(a => a.NetPosition);
            return netPosition;
        }

       private decimal GetNav(List<FundNetAssetValue> navs,DateTime referenceDate)
        {
            var nav = navs.FirstOrDefault(a => a.ReferenceDate == referenceDate);
            var navValue = 0m;
            if (nav != null)
            {
                navValue = nav.MarketValue;

            }
            return navValue;
        }

        public Fund GetFund(FundIds fundId, EntityTypes childEntityType, bool isPrimary)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var fund = context.Funds.Include(a => a.LegalEntity).Include(a => a.Currency.Instrument).FirstOrDefault(a => a.LegalEntityID == (int)fundId);



                bool includeHedging = fund.FundTypeId == (int)FundTypeIds.ShareClass;
                bool IncludeOnlyFX = !isPrimary;

                Fund toReturn = new Fund(fund.LegalEntityID, fund.Name, fund.Currency.IsoCode, fund.CurrencyID, false, fund.IsLongOnly, childEntityType, includeHedging, IncludeOnlyFX, isPrimary);              
                DateTime[] referenceDates = { PreviousReferenceDate, ReferenceDate };
                var navs = context.FundNetAssetValues.Where(a => a.FundId == fund.LegalEntityID && referenceDates.Contains(a.ReferenceDate)).ToList();
                toReturn.Nav = GetNav(navs, ReferenceDate);
                toReturn.PreviousNav = GetNav(navs, PreviousReferenceDate);
                return toReturn;
            }
        }



        




        public List<FXRateDTO> GetFXRates()
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                DateTime[] referenceDates = { PreviousPreviousReferenceDate, PreviousReferenceDate };
                var todayRates = context.FXRates.Include(a => a.FromCurrency.Instrument).Include(a => a.ToCurrency.Instrument)
                    .Where(a => referenceDates.Contains(a.ReferenceDate) && a.ReferenceDate == a.ForwardDate).ToList();
                return todayRates.GroupBy(g => new { FromCurrency = g.FromCurrency.Instrument.Name, ToCurrency = g.ToCurrency.Instrument.Name })
                      .Select(a => {

                          var previousPrevious = a.FirstOrDefault(f => f.ReferenceDate == PreviousPreviousReferenceDate);
                          var previous = a.FirstOrDefault(f => f.ReferenceDate == PreviousReferenceDate);
                          return new FXRateDTO()

                          {
                              FromCurrency = a.Key.FromCurrency,
                              ToCurrency = a.Key.ToCurrency,
                              PreviousPreviousValue = previousPrevious == null ? 1 : previousPrevious.Value,
                              PreviousValue = previous == null ? 1 : previous.Value
                          };                           
                      }).ToList();
            }
        }

        public InstrumentDTO GetInstrument(string ticker)
        {
            using (KeeleyModel context = new KeeleyModel())
            {
                var instrumentMarkets = context.InstrumentMarkets
                    .Include(a => a.Instrument.InstrumentClass.ParentInstrumentClassRelationships)
                    .Include(a => a.Market.LegalEntity.Country)
                    .Include(a => a.PriceCurrency.Instrument)
                    
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


        public List<InstrumentDTO> GetInstruments(List<Identifier> identifiers)
        {
            int[] instrumentMarketIds = identifiers.Where(a => a.Id.HasValue).Select(a => a.Id.Value).Distinct().ToArray();

            string[] tickers = identifiers.Where(a => !a.Id.HasValue).Select(a => a.Code).Distinct().ToArray();

            using (KeeleyModel context = new KeeleyModel())
            {

                var instrumentMarkets = context.InstrumentMarkets
                    .Include(a => a.Instrument.InstrumentClass.ParentInstrumentClassRelationships)
                    .Include(a => a.Market.LegalEntity.Country)
                    .Include(a => a.PriceCurrency.Instrument)
                    .Where(a => instrumentMarketIds.Contains(a.InstrumentMarketID) ).ToList();

                List<InstrumentDTO> dtos = instrumentMarkets.Select(a => InstrumentBuilder.Instance.Get(a)).ToList();

                instrumentMarkets = context.InstrumentMarkets
                    .Include(a => a.Instrument.InstrumentClass.ParentInstrumentClassRelationships)
                    .Include(a => a.Market.LegalEntity.Country)
                    .Include(a => a.PriceCurrency.Instrument)

                    .Where(a => tickers.Contains(a.BloombergTicker) && a.Instrument.InstrumentClassID != (int)InstrumentClassIds.ContractForDifference).ToList();

                var tickerDtos = instrumentMarkets.Select(a => InstrumentBuilder.Instance.Get(a)).ToList();

                dtos.AddRange(tickerDtos);
                return dtos;
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
                    string exchangeCode = ticker.Substring(ticker.IndexOf(' ')+1);
                    exchangeCode = exchangeCode.Substring(0, exchangeCode.IndexOf(' '));
                    var market = context.Markets.Include(a=>a.LegalEntity.Country).FirstOrDefault(a => a.BBExchangeCode == exchangeCode);
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
