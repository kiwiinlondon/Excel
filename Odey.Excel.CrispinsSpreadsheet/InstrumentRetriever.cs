using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class InstrumentRetriever
    {
        public InstrumentRetriever(BloombergSecuritySetup bloombergSecuritySetup, DataAccess dataAccess)
        {
            _bloombergSecuritySetup = bloombergSecuritySetup;
            _dataAccess = dataAccess;
        }

        BloombergSecuritySetup _bloombergSecuritySetup;
        DataAccess _dataAccess;

        public InstrumentDTO Get(string ticker, out string message)
        {
            message = null;
            InstrumentDTO dto = _dataAccess.GetInstrument(ticker);
            if (dto==null)
            {
                if (_bloombergSecuritySetup.ValidateTicker(ticker, out message))
                {
                    dto = _bloombergSecuritySetup.GetInstrument(ticker);
                    if (dto==null)
                    {
                        message = $"Unable To Find Ticker {message} in Bloomberg";
                    }
                    else
                    {
                        EnhanceBloombergDTO(ticker,dto);
                    }
                }               
            }
            return dto;
        }

        private void EnhanceBloombergDTO(string ticker, InstrumentDTO dto)
        {

            dto.Identifier = new Identifier(null,ticker);
            dto.InstrumentTypeId = InstrumentTypeIds.Normal;
            dto.PriceDivisor = 1;
            dto.AssetClass = EntityBuilder.EquityLabel;
            _dataAccess.AddExchangeCountryToInstrument(dto);
        }

        public string FixTicker(string ticker)
        {            
            if (!ticker.ToUpper().EndsWith(" EQUITY") && !ticker.ToUpper().EndsWith(" GOVT") && !ticker.ToUpper().EndsWith(" COMDTY") && !ticker.ToUpper().EndsWith(" INDEX"))
            {
                ticker = $"{ticker} Equity";
            }
            ticker = Regex.Replace(ticker, @"\s+", " ");

            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex(@"[ ]{2,}", options);
            ticker = regex.Replace(ticker, @" ");
            ticker = ticker
                .Replace(" AT ", " AU ")
                .Replace(" UN ", " US ")
                .Replace(" UE ", " US ")
                .Replace(" UA ", " US ")
                .Replace(" UQ ", " US ")
                .Replace(" UW ", " US ")
                .Replace(" UC ", " US ")
                .Replace(" UP ", " US ")
                .Replace(" UU ", " US ")
                .Replace(" UR ", " US ")
                .Replace(" GF ", " GY ")
                .Replace(" GB ", " GY ")
                .Replace(" GT ", " GY ")
                .Replace(" GR ", " GY ")
                .Replace(" CT ", " CN ")
                .Replace(" CV ", " CN ")
                .Replace(" KP ", " KS ")
                .Replace(" SE ", " SW ")
                .Replace(" VX ", " SW ")
                .Replace(" JQ ", " JT ")
                .Replace(" JP ", " JT ")
                .Replace(" AN ", " AU ")
                .Replace(" SM ", " SQ ")
                .Replace(" KP ", " KS ")
                //.Replace(" BS ", " BZ ")
                .Replace(" BZ ", " BS ")
                .Replace(" NI ", " IM ")
                .Replace(" RU ", " RM ")//Sberbank had no RU record so changed to use RM
                .Replace(" RR ", " RM ")
                //.Replace(" NS ", " NO ")//NS is a valid exchange
                .Replace(" NP ", " FP ")
                .Replace(" IB ", " IN ")
                .Replace(" CD Equity", " CP Equity");
            return ticker;
        }

        public bool ValidateTicker(string ticker, out string message)
        {
            message = null;
            int countOfSpaces = ticker.Count(a => a == ' ');

            if (ticker.ToUpper().EndsWith("EQUITY"))
            {
                if (countOfSpaces != 2)
                {
                    message = "No Exchange on Ticker";
                    return false;
                }
            }
            else if (countOfSpaces != 1)
            {
                message = "Incorrect Format on ticker";
                return false;
            }

            

            return true;
        }


    }
}
