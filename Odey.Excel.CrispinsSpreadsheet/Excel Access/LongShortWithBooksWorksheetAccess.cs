using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class LongShortWithBooksWorksheetAccess : WorksheetAccess
    {
        public LongShortWithBooksWorksheetAccess(Worksheet worksheet) : base(worksheet)
        {
        }
        protected override string ContributionBookColumn => "O";
        protected override string ContributionFundColumn => "P";
        protected override string ExposureColumn => "Q";
        protected override string ExposurePercentageBookColumn => "R";
        protected override string ExposurePercentageFundColumn => "S";
        protected override string ShortBookColumn => "T";
        protected override string LongBookColumn => "U";
        protected override string ShortFundColumn => null;
        protected override string LongFundColumn => null;
        protected override string PriceMultiplierColumn => "V";
        protected override string InstrumentTypeColumn => "W";
        protected override string PriceDivisorColumn => "X";
        protected override string ShortBookWinnersColumn => "Y";
        protected override string LongBookWinnersColumn => "Z";
        protected override string ShortFundWinnersColumn => null;
        protected override string LongFundWinnersColumn => null;
        protected override string NavColumn => "AA";
        protected override string PreviousClosePriceColumn => "AB";
        protected override string PreviousPriceChangeColumn => "AC";
        protected override string PreviousPricePercentageChangeColumn => "AD";
        protected override string PreviousNetPositionColumn => "AE";
        protected override string PreviousFXRateColumn => "AF";
        protected override string PreviousContributionBookColumn => "AG";
        protected override string PreviousContributionFundColumn => "AH";       
        protected override string PreviousNavColumn => "AI";
    }
}
