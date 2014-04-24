using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;

namespace TimesheetEventHandler
{
    // Define a custom logging service for the ULS log.
    public class LoggingService : SPDiagnosticsServiceBase
    {
        private const string LOG_SERVICE_NAME = "Project Test Logging Service";
        private const string PRODUCT_DIAGNOSTIC_NAME = "Project Server Event Handler";
        private const uint EVENT_ID = 5050;

        // ULS categories:
        public const string PROJECT_INFO = "Event Handler Information";
        public const string PROJECT_WARNING = "Event Handler Warning";

        private static LoggingService activeLoggingService;

        private LoggingService() : base(LOG_SERVICE_NAME, SPFarm.Local)
        {
        }

        // Register the product name and set the ULS log categories that are available.
        protected override IEnumerable<SPDiagnosticsArea> ProvideAreas()
        {
            List<SPDiagnosticsArea> areas = new List<SPDiagnosticsArea>
            {
                new SPDiagnosticsArea(
                    PRODUCT_DIAGNOSTIC_NAME,
                    new List<SPDiagnosticsCategory>
                    {
                        new SPDiagnosticsCategory(PROJECT_INFO,
                                                  TraceSeverity.Verbose, 
                                                  EventSeverity.Information),
                        new SPDiagnosticsCategory(PROJECT_WARNING,
                                                  TraceSeverity.Unexpected,
                                                  EventSeverity.Warning),
                    })
            };
            return areas;
        }

        // Create a LoggingService instance.
        public static LoggingService Active
        {
            get
            {
                if (activeLoggingService == null)
                    activeLoggingService = new LoggingService();
                return activeLoggingService;
            }
        }

        // Write an information message to the ULS log.
        public static void LogMessage(string categoryName, string message)
        {
            SPDiagnosticsCategory category =
                LoggingService.Active.Areas[PRODUCT_DIAGNOSTIC_NAME].Categories[categoryName];
            LoggingService.Active.WriteTrace(EVENT_ID, category, TraceSeverity.Verbose, message);
        }

        // Write an error message to the ULS log.
        public static void LogError(string categoryName, string message)
        {
            SPDiagnosticsCategory category =
                LoggingService.Active.Areas[PRODUCT_DIAGNOSTIC_NAME].Categories[categoryName];
            LoggingService.Active.WriteTrace(EVENT_ID, category, TraceSeverity.Unexpected, message);
        }
    }
}
