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
            _sheetAccess.DisableCalculations();
            decimal nav;
            decimal previousNav;
            var dTOs = _dataAccess.Get(fundId, referenceDate, out nav, out previousNav);

            _sheetAccess.WriteNAVs(previousNav,nav);
            Dictionary<string, CountryLocation> countries = _sheetAccess.GetCountries();
            AddDTOsToCountrys(dTOs, countries);
            WriteCountries(countries);
            _sheetAccess.EnableCalculations();
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
            
            if (!country.TotalRowNumber.HasValue)
            {
                _sheetAccess.AddCountryTotalRow(lastRow, country);
            }

            int currentRow = country.TotalRowNumber.Value;
            foreach (var location in country.TickerRows.Values.OrderByDescending(a=>a.Name))
            {
                if (location.RowNumber.HasValue)
                {
                    _sheetAccess.UpdateTickerRow(true,location);

                    currentRow = location.RowNumber.Value;
                }
                else
                {
                    _sheetAccess.AddTickerRow(location, currentRow);
                    country.TotalRowNumber++;
                }
            }
            _sheetAccess.UpdateSums(country);
            return country.FirstRowNumber.Value;
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

                    location.PreviousNetPosition = dto.PreviousNetPosition;
                    location.NetPosition = dto.CurrentNetPosition;                    
                    location.Currency = dto.Currency;
                    location.Name = dto.Name;
                    location.PriceDivisor = dto.PriceDivisor;
                    location.TickerTypeId = dto.TickerTypeId;
                    location.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
                    location.OdeyPreviousPrice = dto.PreviousPrice;
                    location.OdeyCurrentPrice = dto.CurrentPrice;
                }
                else
                {
                    location = new Location(null, dto.Ticker, dto.Name, dto.PreviousNetPosition, dto.CurrentNetPosition, dto.TickerTypeId, dto.PreviousPreviousPrice, dto.PreviousPrice, dto.CurrentPrice, dto.Currency, dto.PriceDivisor, null);
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
