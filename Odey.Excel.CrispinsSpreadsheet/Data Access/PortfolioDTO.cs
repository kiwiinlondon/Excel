using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class PortfolioDTO
    {

        public PortfolioDTO(string book,InstrumentDTO instrument,
             decimal previousNetPosition, decimal currentNetPosition, decimal? previousPreviousPrice, decimal? previousPrice, decimal? currentPrice, bool previousPreviousPriceIsManual, bool previousPriceIsManual, bool currentPriceIsManual)
        {

            Book = book;
            Instrument = instrument;
            
            PreviousNetPosition = previousNetPosition;
            CurrentNetPosition = currentNetPosition;
            
            PreviousPreviousPrice = previousPreviousPrice;
            PreviousPrice = previousPrice;
            CurrentPrice = currentPrice;

            CurrentPriceIsManual = currentPriceIsManual;

            PreviousPriceIsManual = previousPriceIsManual;

            PreviousPreviousPriceIsManual = previousPreviousPriceIsManual;

        }

        public string Book { get; set; }

        public InstrumentDTO Instrument { get; set;}

        public decimal CurrentNetPosition { get; set; }
        public decimal PreviousNetPosition { get; set; }

        public decimal? CurrentPrice { get; set; }

        public decimal? PreviousPrice { get; set; }

        public decimal? PreviousPreviousPrice { get; set; }

        public bool CurrentPriceIsManual { get; set; }

        public bool PreviousPriceIsManual { get; set; }

        public bool PreviousPreviousPriceIsManual { get; set; }
    }
}
