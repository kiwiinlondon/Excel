using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class LongShortWithoutBooksWorksheetAccess : WorksheetAccess
    {
        public LongShortWithoutBooksWorksheetAccess(Worksheet worksheet) : base(worksheet)
        {
        }
        protected override string ContributionBookColumn => null;
        protected override string ContributionFundColumn => "O";
        protected override string ExposureColumn => "P";
        protected override string ExposurePercentageBookColumn => null;
        protected override string ExposurePercentageFundColumn => "Q";
        protected override string ShortBookColumn => null;
        protected override string LongBookColumn => null;
        protected override string ShortFundColumn => "R";
        protected override string LongFundColumn => "S";
        protected override string PriceMultiplierColumn => "T";
        protected override string InstrumentTypeColumn => "U";
        protected override string PriceDivisorColumn => "V";
        protected override string ShortBookWinnersColumn => null;
        protected override string LongBookWinnersColumn => null;

        protected override string ShortFundWinnersColumn => "W";
        protected override string LongFundWinnersColumn => "X";
        protected override string NavColumn => "Y";
        protected override string PreviousClosePriceColumn => "Z";
        protected override string PreviousPriceChangeColumn => "AA";
        protected override string PreviousPricePercentageChangeColumn => "AB";
        protected override string PreviousNetPositionColumn => "AC";
        protected override string PreviousFXRateColumn => "AD";
        protected override string PreviousContributionBookColumn => null;
        protected override string PreviousContributionFundColumn => "AE";
        protected override string PreviousNavColumn => "AF";
    }
}
