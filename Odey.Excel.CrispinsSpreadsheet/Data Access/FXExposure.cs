
using Odey.Framework.Keeley.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    internal class FXExposure
    {
        public FXExposure(Currency currency, int bookId, InstrumentMarket instrumentMarket, string label, FXExposureTypeIds fxExposureTypeId, DateTime referenceDate, decimal netPosition, decimal price, decimal marketValue,decimal fundNav)
        {
            BookId = bookId;
            Currency = currency;
            InstrumentMarket = instrumentMarket;
            Label = label;
            ReferenceDate = referenceDate;
            NetPosition = netPosition;
            Price = price;
            MarketValue = marketValue;
            FXExposureTypeId = fxExposureTypeId;
            FundNav = fundNav;
        }

        public Currency Currency { get; private set; }
        public int CurrencyId => Currency.InstrumentID;

        public int BookId { get; private set; }       

        public InstrumentMarket InstrumentMarket { get; private set; }
        public int InstrumentMarketID => InstrumentMarket.InstrumentMarketID;

        public string Label { get; private set; }

        public DateTime ReferenceDate { get; private set; }
        public decimal NetPosition { get; private set; }
        public decimal Price { get; private set; }

        public decimal MarketValue { get; private set; }

        public FXExposureTypeIds FXExposureTypeId { get; private set; }

        public decimal FundNav { get; private set; }

        public override string ToString()
        {
            return $"{Currency.IsoCode} {Label} {Math.Round(MarketValue/FundNav*100,2)}%";
        }
    }
}
