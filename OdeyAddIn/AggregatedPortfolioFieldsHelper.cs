using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OdeyAddIn
{
    public static class AggregatedPortfolioFieldsHelper
    {
        public static AggregatedPortfolioFields[] Get(bool returnRawData, AggregatedPortfolioOutputOptions outputOption)
        {
            List<AggregatedPortfolioFields> aggregatedPortfolioFields = new List<AggregatedPortfolioFields>();
            switch (outputOption)
            {
                case AggregatedPortfolioOutputOptions.Gross:
                    aggregatedPortfolioFields.Add(AggregatedPortfolioFields.GrossPercentNav);
                    if (returnRawData)
                    {
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.Gross);
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.FundNav);
                    }
                    break;
                case AggregatedPortfolioOutputOptions.LongShort:
                    aggregatedPortfolioFields.Add(AggregatedPortfolioFields.ShortPercentNav);
                    aggregatedPortfolioFields.Add(AggregatedPortfolioFields.LongPercentNav);
                    if (returnRawData)
                    {
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.Short);
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.Long);
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.FundNav);
                    }
                    break;
                case AggregatedPortfolioOutputOptions.Net:
                    aggregatedPortfolioFields.Add(AggregatedPortfolioFields.NetPercentNav);
                    if (returnRawData)
                    {
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.Net);
                        aggregatedPortfolioFields.Add(AggregatedPortfolioFields.FundNav);
                    }
                    break;
                default:
                    throw new ApplicationException(String.Format("Unknown Output Option {0}", outputOption));
            }
            return aggregatedPortfolioFields.ToArray();
        }
    }
}
