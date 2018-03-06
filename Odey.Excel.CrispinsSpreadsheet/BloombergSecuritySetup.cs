using Bloomberglp.Blpapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class BloombergSecuritySetup
    {
        private static readonly Name EXCEPTIONS = new Name("exceptions");
        private static readonly Name FIELD_ID = new Name("fieldId");
        private static readonly Name REASON = new Name("reason");
        private static readonly Name CATEGORY = new Name("category");
        private static readonly Name DESCRIPTION = new Name("description");
        private static readonly Name ERROR_CODE = new Name("errorCode");
        private static readonly Name SOURCE = new Name("source");
        private static readonly Name SECURITY_ERROR = new Name("securityError");
        private static readonly Name MESSAGE = new Name("message");
        private static readonly Name RESPONSE_ERROR = new Name("responseError");
        private static readonly Name SECURITY_DATA = new Name("securityData");
        private static readonly Name FIELD_EXCEPTIONS = new Name("fieldExceptions");
        private static readonly Name ERROR_INFO = new Name("errorInfo");
        private static readonly Name SECURITY_NAME = new Name("SECURITY_NAME");
        private Session d_session;

        private SessionOptions d_sessionOptions;

        private bool createSession()
        {
            if (d_session != null)
            {
                // Session.Stop needs to be called asynchronously to 
                // prevent blocking, while waiting for GUI event processing 
                // to return.
                d_session.Stop(AbstractSession.StopOption.ASYNC);
            }


            string serverHost = "localhost";
            int serverPort = 8194;

            // set sesson options
            d_sessionOptions = new SessionOptions();
            d_sessionOptions.ServerHost = serverHost;
            d_sessionOptions.ServerPort = serverPort;

            // create asynchronous session
            d_session = new Session(d_sessionOptions);

            return d_session.Start();
        }

        public InstrumentDTO GetInstrument(string ticker)
        {

            if (!createSession())
            {
                throw new ApplicationException("Failed to start Bloomberg Session");
            }
            // open reference data service
            if (!d_session.OpenService("//blp/refdata"))
            {
                throw new ApplicationException("Failed to open //blp/refdata");
            }
            Service refDataService = d_session.GetService("//blp/refdata");
            Request request = refDataService.CreateRequest("ReferenceDataRequest");

            Element securities = request.GetElement("securities");
            Element fields = request.GetElement("fields");
            Element requestOverrides = request.GetElement("overrides");
            request.Set("returnEids", true);
            securities.AppendValue(ticker);
            fields.AppendValue("SECURITY_NAME");

            CorrelationID cID = new CorrelationID(1);
            d_session.Cancel(cID);
            // send request
            d_session.SendRequest(request, cID);

            InstrumentDTO dto = null;
            while (true)
            {
                Event eventObj = d_session.NextEvent();
                // process data
                dto = processEvent(eventObj, d_session);
                if (eventObj.Type == Event.EventType.RESPONSE)
                {
                    break;
                }
            }
            return dto;
        }


        public bool ValidateTicker(string ticker, out string message)
        {
            message = null;
            if (!ticker.ToUpper().EndsWith("EQUITY"))
            {
                message = "Only Equities allowed at the moment";
                return false;
            }
            return true;
        }

        private InstrumentDTO processEvent(Event eventObj, Session session)
        {
            InstrumentDTO dto = null;
            switch (eventObj.Type)
            {
                case Event.EventType.RESPONSE:
                    // process final respose for request
                    dto = processRequestDataEvent(eventObj, session);
                    break;
                case Event.EventType.PARTIAL_RESPONSE:
                    // process partial response
                    dto = processRequestDataEvent(eventObj, session);
                    break;
                default:
                    processMiscEvents(eventObj, session);
                    break;
            }
            return dto;

        }

        private InstrumentDTO processRequestDataEvent(Event eventObj, Session session)
        {
            InstrumentDTO dto = null;
            // process message
            foreach (Message msg in eventObj)
            {
                // get message correlation id
                int cId = (int)msg.CorrelationID.Value;
                if (msg.MessageType.Equals(Bloomberglp.Blpapi.Name.GetName("ReferenceDataResponse")))
                {
                    // process security data
                    Element secDataArray = msg.GetElement(SECURITY_DATA);
                    int numberOfSecurities = secDataArray.NumValues;
                    for (int index = 0; index < numberOfSecurities; index++)
                    {
                        Element secData = secDataArray.GetValueAsElement(index);
                        Element fieldData = secData.GetElement("fieldData");
                        // get security index
                        int rowIndex = secData.GetElementAsInt32("sequenceNumber");

                        // check for field error
                        if (secData.HasElement(FIELD_EXCEPTIONS))
                        {
                            string message = "";
                            // process error
                            Element error = secData.GetElement(FIELD_EXCEPTIONS);
                            if (error.NumValues > 0)
                            {
                                for (int errorIndex = 0; errorIndex < error.NumValues; errorIndex++)
                                {
                                    Element errorException = error.GetValueAsElement(errorIndex);
                                    string field = errorException.GetElementAsString(FIELD_ID);
                                    Element errorInfo = errorException.GetElement(ERROR_INFO);
                                    message += errorInfo.GetElementAsString(MESSAGE);

                                }
                                throw new ApplicationException(message);
                            }
                        }
                        // check for security error
                        if (secData.HasElement(SECURITY_ERROR))
                        {
                            Element error = secData.GetElement(SECURITY_ERROR);
                            string errorMessage = error.GetElementAsString(MESSAGE);
                            throw new ApplicationException(errorMessage);
                        }
                        // process data
                 
                        String dataValue = string.Empty;
                        if (fieldData.HasElement(SECURITY_NAME))
                        {
                            Element item = fieldData.GetElement(SECURITY_NAME);
                            dataValue =  item.GetValueAsString();                            
                        }
                        dto = new InstrumentDTO() { Name = dataValue };
                    }
                }
            }
            return dto;
        }


        private void processMiscEvents(Event eventObj, Session session)
        {
            foreach (Message msg in eventObj)
            {
                switch (msg.MessageType.ToString())
                {
                    case "SessionStarted":
                        // "Session Started"
                        break;
                    case "SessionTerminated":
                    case "SessionStopped":
                        // "Session Terminated"
                        break;
                    case "ServiceOpened":
                        // "Reference Service Opened"
                        break;
                    case "RequestFailure":
                        Element reason = msg.GetElement(REASON);
                        string message = string.Concat("Error: Source-", reason.GetElementAsString(SOURCE),
                            ", Code-", reason.GetElementAsString(ERROR_CODE), ", category-", reason.GetElementAsString(CATEGORY),
                            ", desc-", reason.GetElementAsString(DESCRIPTION));
                        throw new ApplicationException($"Bloomberg Failure: {message}");
                    default:
                        break;
                }
            }
        }
    }
}
