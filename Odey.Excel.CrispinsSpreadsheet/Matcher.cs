using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Matcher
    {
        private DataAccess _dataAccess;
        private SheetAccess  _sheetAccess;
        public Matcher (DataAccess dataAccess,SheetAccess sheetAccess)
        {
            _dataAccess = dataAccess;
            _sheetAccess = sheetAccess;
        }
            
        public void Match(int fundId, DateTime referenceDate)
        {
            decimal nav;
            var dTOs = _dataAccess.Get(fundId, referenceDate, out nav);//.Where(a => a.Ticker == "SKY LN Equity" || a.Ticker == "INTU LN Equity" || a.Ticker == "MTS AU Equity" || a.Ticker == "MON US Equity" || a.Ticker == "8591 JT Equity" || a.Ticker == "TALK LN Equity" || a.Ticker == "WFT US Equity").ToList();

            _sheetAccess.WriteNAV(nav);
            Dictionary<string, CountryLocation> countries = _sheetAccess.GetCountries();
            AddDTOsToCountrys(dTOs, countries);
            WriteCountries(countries);
        }

        private void AddDTOsToCountrys(List<PortfolioDTO> dTOs, Dictionary<string, CountryLocation> countries)
        {
            foreach (PortfolioDTO dTO in dTOs)
            {
                CountryLocation country;
                if (!countries.TryGetValue(dTO.CountryIsoCode, out country))
                {
                    country = new CountryLocation()
                    {
                        IsoCode = dTO.CountryIsoCode,
                        Name = dTO.CountryName
                    };
                    countries.Add(country.IsoCode,country);
                }
                country.PortfolioDtos.Add(dTO);
            }
        }

        private void WriteCountries(Dictionary<string, CountryLocation> countries)
        {
            int lastTotalRow = _sheetAccess.LastRow;
            foreach (var country in countries.Values.OrderByDescending(a=>a.Name))
            {
                lastTotalRow= WriteCountry(lastTotalRow,country);
            }
        }

        private int WriteCountry(int lastRow,CountryLocation country)
        {
            MatchDTOToTickerRows(country);
            
            if (!country.TotalRow.HasValue)
            {
                _sheetAccess.AddCountryTotalRow(lastRow, country);
            }

            int currentRow = country.TotalRow.Value;
            foreach (var location in country.TickerRows.Values.OrderByDescending(a=>a.Name))
            {
                if (location.Row.HasValue)
                {
                    _sheetAccess.UpdateTickerRow(true,location);

                    currentRow = location.Row.Value;
                }
                else
                {
                    _sheetAccess.AddTickerRow(location, currentRow);
                    country.TotalRow++;
                }
            }
            _sheetAccess.UpdateSums(country);
            return country.FirstRow.Value;
        }

        private void MatchDTOToTickerRows(CountryLocation country)
        {
            List<string> tickersToZeroOut = country.TickerRows.Keys.ToList();
            foreach (PortfolioDTO dto in country.PortfolioDtos)
            {
                Location location;
                if (country.TickerRows.TryGetValue(dto.Ticker, out location))
                {
                    tickersToZeroOut.Remove(dto.Ticker);
                    if (location.NetPosition != dto.NetPosition)
                    {
                        location.NetPosition = dto.NetPosition;
                    }
                }
                else
                {
                    location = new Location(null, dto.Ticker, dto.Name, dto.NetPosition, dto.TickerTypeId,dto.Price, dto.Currency, dto.PriceDivisor);
                    country.TickerRows.Add(location.Ticker, location);
                }
            }
            foreach (string ticker in tickersToZeroOut)
            {
                var row = country.TickerRows[ticker];
                row.NetPosition = 0;
            }
        }
    }
}
