﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class LongOnlyWorksheetAccess : WorksheetAccess
    {
        public LongOnlyWorksheetAccess(Worksheet worksheet) : base(worksheet)
        {


        }
        protected override string ContributionBookColumn => null;
        protected override string ContributionFundColumn => "O";
        protected override string ExposureColumn => "P";
        protected override string ExposurePercentageBookColumn => null;
        protected override string ExposurePercentageFundColumn => "Q";
        protected override string ShortBookColumn => null;
        protected override string LongBookColumn => null;
        protected override string PriceMultiplierColumn => "R";
        protected override string InstrumentTypeColumn => "S";
        protected override string PriceDivisorColumn => "T";
        protected override string ShortBookWinnersColumn => null;
        protected override string LongBookWinnersColumn => null;
        protected override string NavColumn => "U";
        protected override string PreviousClosePriceColumn => "V";
        protected override string PreviousPriceChangeColumn => "W";
        protected override string PreviousPricePercentageChangeColumn => "X";
        protected override string PreviousNetPositionColumn => "Y";
        protected override string PreviousFXRateColumn => "Z";
        protected override string PreviousContributionBookColumn => null;
        protected override string PreviousContributionFundColumn => "AA";
        protected override string PreviousNavColumn => "AB";
    }
}
