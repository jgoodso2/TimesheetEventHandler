using System;
using System.Linq;
using System.Diagnostics;
using System.ServiceModel;
using System.Security.Principal;
using Microsoft.SharePoint;
using Microsoft.Office.Project.Server.Events;
using SvcTimeSheet;
using PSLib = Microsoft.Office.Project.Server.Library;
using System.Globalization;
using WCFHelpers;
using System.Data;

namespace TimesheetEventHandler
{
    public class TimesheetEventHandler : TimesheetEventReceiver
    {
        private SvcQueueSystem.QueueSystemClient queueClient;

        private const string EVENT_SOURCE = "Timesheet Event Handler";
        private const int EVENT_ID = 5050;
        private EventLog eventLog = new EventLog();
       
        SvcTimeSheet.TimeSheetClient timesheetClient;
        SvcAdmin.AdminClient adminClient;
        SvcResource.ResourceClient resourceClient;
        public Guid GetResourceUidFromNtAccount(String ntAccount)
        {

            //ntAccount = "i:0#.w|" + ntAccount;  //this is an inconsequential change
            ntAccount = ntAccount.Trim('\"');

            string ntAccountCopy = ntAccount;
            
            SvcResource.ResourceDataSet rds = new SvcResource.ResourceDataSet();

            Microsoft.Office.Project.Server.Library.Filter filter = new Microsoft.Office.Project.Server.Library.Filter();
            filter.FilterTableName = rds.Resources.TableName;


            Microsoft.Office.Project.Server.Library.Filter.Field ntAccountField1 = new Microsoft.Office.Project.Server.Library.Filter.Field(rds.Resources.TableName, rds.Resources.WRES_ACCOUNTColumn.ColumnName);
            filter.Fields.Add(ntAccountField1);

            Microsoft.Office.Project.Server.Library.Filter.Field ntAccountField2 = new Microsoft.Office.Project.Server.Library.Filter.Field(rds.Resources.TableName, rds.Resources.RES_IS_WINDOWS_USERColumn.ColumnName);
            filter.Fields.Add(ntAccountField2);

            Microsoft.Office.Project.Server.Library.Filter.FieldOperator op = new Microsoft.Office.Project.Server.Library.Filter.FieldOperator(Microsoft.Office.Project.Server.Library.Filter.FieldOperationType.Equal,
                rds.Resources.WRES_ACCOUNTColumn.ColumnName, ntAccountCopy);
            filter.Criteria = op;



            rds = resourceClient.ReadResources(filter.GetXml(), false);

            var obj = (Guid)rds.Resources.Rows[0]["RES_UID"];
            return obj;
        }
        public void SetImpersonation(Guid resourceGuid)
        {
            Guid trackingGuid = Guid.NewGuid();


            bool isWindowsUser = true;
            Guid siteId = Guid.Empty;           // Project Web App site ID.
            CultureInfo languageCulture = null; // The language culture is not used.
            CultureInfo localeCulture = null;   // The locale culture is not used.


            WcfHelpers.SetImpersonationContext(isWindowsUser,

                resourceClient.ReadResource(resourceGuid).Resources[0].RES_NAME, resourceGuid, trackingGuid, siteId,
                                               languageCulture, localeCulture);
            WCFHelpers.WcfHelpers.UseCorrectHeaders(true);
        }
        private void SaveTimesheet(Guid userId, SvcTimeSheet.TimesheetDataSet tsDs, Guid tsGuid)
        {

            try
            {
                Guid jobGuid = Guid.NewGuid();
               

                using (OperationContextScope scope = new OperationContextScope(timesheetClient.InnerChannel))
                {
                    SetImpersonation(userId);
                    var temp = tsDs.GetChanges();
                    timesheetClient.QueueUpdateTimesheet(jobGuid,
                         tsGuid,
                        (SvcTimeSheet.TimesheetDataSet)tsDs);  //Saves the specified timesheet data to the Published database
                }
                bool res = QueueHelper.WaitForQueueJobCompletion(jobGuid, (int)SvcQueueSystem.QueueMsgType.TimesheetUpdate, queueClient);
                if (!res) throw new Exception();
            }
            catch (Exception tex) { throw new Exception(); }
        }

        


        public override void OnSubmitted(PSLib.PSContextInfo contextInfo, TimesheetPostEventArgs e)
        {
            
            try
            {
                base.OnSubmitted(contextInfo, e);
                SetClientEndpoint(contextInfo.SiteGuid);
                Guid pwaGuid = contextInfo.SiteGuid;
                using (OperationContextScope scope = new OperationContextScope(timesheetClient.InnerChannel))
                {
                    var timesheetCopy = timesheetClient.ReadTimesheet(e.TsUID);
                    var timesheet = timesheetCopy;
                    SetImpersonation(contextInfo.UserGuid);
                    //Get current period ID
                    var currentGuid = timesheet.Headers[0].WPRD_UID;
                    //Read all Timesheet Periods
                    var periods = adminClient.ReadPeriods(SvcAdmin.PeriodState.All).TimePeriods.OrderBy(t => t.WPRD_START_DATE).ToList();
                    //Find the index of next period
                    int index = periods.FindIndex(t => t.WPRD_UID == currentGuid);
                    var nextPeriod = (index == periods.Count() - 1) ? periods.ElementAt(index) : periods.ElementAt(index + 1);
                    //If the current period id is not the last period id
                    if (periods[index].WPRD_UID != nextPeriod.WPRD_UID)
                    {
                        //Read timesheet for next period
                        var nextTimesheet = timesheetClient.ReadTimesheetByPeriod(contextInfo.UserGuid, nextPeriod.WPRD_UID, SvcTimeSheet.Navigation.Current);
                        
                        // If next timesheet is not yet created only then
                        if (nextTimesheet.Headers.Count == 0)
                        {
                            Guid TSUID = Guid.Empty;
                            //no prepopulationa
                            CreateTimesheet(contextInfo.UserGuid, nextPeriod.WPRD_UID, ref TSUID, ref nextTimesheet);
                            //for each line that was present in the current timesheet
                            nextTimesheet.Lines.Clear();
                            foreach (var line in timesheet.Lines)
                            {
                               
                                //add the line to the next timesheet
                                try
                                {
                                   
                                
                                var lineRow = nextTimesheet.Lines.AddLinesRow(Guid.NewGuid(), nextTimesheet.Headers[0], line.ASSN_UID, line.TASK_UID, line.PROJ_UID,
                                    line.TS_LINE_CLASS_UID, line.TS_LINE_COMMENT, line.TS_LINE_VALIDATION_TYPE, line.TS_LINE_CACHED_ASSIGN_NAME,
                                   line.TS_LINE_CACHED_PROJ_NAME, line.TS_LINE_CACHED_PROJ_REVISION_COUNTER, line.TS_LINE_CACHED_PROJ_REVISION_RANK, line.TS_LINE_IS_CACHED, 0,
                                   line.TS_LINE_STATUS,
                                   0, line.TS_LINE_TASK_HIERARCHY);
                                var date = nextPeriod.WPRD_START_DATE;
                                    if(nextTimesheet.Lines.Any(t=>t.ASSN_UID == lineRow.ASSN_UID))
                                    {
                                        continue;
                                    }
                                Guid[] uids = new Guid[] { lineRow.TS_LINE_UID };
                                timesheetClient.PrepareTimesheetLine(TSUID, ref nextTimesheet, uids);
                                var actuals = lineRow.GetActualsRows();
                                foreach (var actual in actuals)
                                {
                                    actual.SetTS_ACT_NON_BILLABLE_OVT_VALUENull();
                                    actual.SetTS_ACT_NON_BILLABLE_VALUENull();
                                    actual.SetTS_ACT_OVT_VALUENull();
                                    actual.SetTS_ACT_VALUENull();
                                }
                                }
                                catch (Exception)
                                {
                                    continue;
                                }
                            }
                            //Save next timesheet
                            SaveTimesheet(contextInfo.UserGuid, nextTimesheet, TSUID);

                        }
                    }


                }
                WriteLogEntries(contextInfo.UserName, e.TsUID, "Successfully done updating next timesheet");
            }
            catch (Exception ex)
            {
                WriteLogEntries(contextInfo.UserName, e.TsUID, ex.Message);
            }

        }


        private void CreateTimesheet(Guid userUid, Guid periodUID, ref Guid tuid, ref SvcTimeSheet.TimesheetDataSet tsDs)
        {
            tsDs = new SvcTimeSheet.TimesheetDataSet();
            SvcTimeSheet.TimesheetDataSet.HeadersRow headersRow = tsDs.Headers.NewHeadersRow();
            headersRow.RES_UID = userUid;  // cant be null.
            tuid = Guid.NewGuid();
            headersRow.TS_UID = tuid;
            headersRow.WPRD_UID = periodUID;
            headersRow.TS_NAME = "Timesheet";
            headersRow.TS_COMMENTS = "Timesheet Created via custom Prepopulation";
            headersRow.TS_ENTRY_MODE_ENUM = (byte)PSLib.TimesheetEnum.EntryMode.Daily;
            tsDs.Headers.AddHeadersRow(headersRow);

            using (OperationContextScope scope = new OperationContextScope(timesheetClient.InnerChannel))
            {
                SetImpersonation(userUid);
                timesheetClient.CreateTimesheet(tsDs, SvcTimeSheet.PreloadType.None);
            }
            using (OperationContextScope scope = new OperationContextScope(timesheetClient.InnerChannel))
            {
                SetImpersonation(userUid);
                tsDs = timesheetClient.ReadTimesheet(tuid); //calling ReadTimesheet to pre populate with default server settings
            }
        }


        // Write entries to the event log and to the ULS log.
        private void WriteLogEntries(
            string userName, Guid tsUID,string error)
        {
            try
            {
                EventLogEntryType entryType = EventLogEntryType.Error;
                string taskInfo = "";

                eventLog.Source = EVENT_SOURCE;

                string logEntry = "User: " + userName;
                logEntry += "\nTSUID: " + tsUID.ToString();

                logEntry += "Error:" + error;
                // Create an event log entry.
                eventLog.WriteEntry(logEntry, entryType, EVENT_ID);

                // Create a ULS log entry.


                LoggingService.LogError(LoggingService.PROJECT_WARNING, logEntry);

            }
            catch (Exception)
            {
                
                
            }
            
        }

        // Programmatically set the WCF endpoint for the LookupTable client.
        private void SetClientEndpoint(Guid pwaUid)
        {
            const int MAXSIZE = 500000000;
            const string svcRouter = "/_vti_bin/PSI/ProjectServer.svc";

            BasicHttpBinding binding = null;
            //TODO: look for dispose issue
            using (SPSite pwaSite = new SPSite(pwaUid))
            {
                string pwaUrl = pwaSite.Url;

                if (pwaUrl.Contains("https:"))
                {
                    // Create a binding for HTTPS.
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
                }
                else
                {
                    // Create a binding for HTTP.
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
                }

                binding.Name = "basicHttpConf";
                binding.SendTimeout = TimeSpan.MaxValue;
                binding.MaxReceivedMessageSize = MAXSIZE;
                binding.ReaderQuotas.MaxNameTableCharCount = MAXSIZE;
                binding.MessageEncoding = WSMessageEncoding.Text;
                binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;

                // The endpoint address is the ProjectServer.svc router for all public PSI calls.
                EndpointAddress address = new EndpointAddress(pwaUrl + svcRouter);

                timesheetClient = new SvcTimeSheet.TimeSheetClient(binding, address);
                timesheetClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                timesheetClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                adminClient = new SvcAdmin.AdminClient(binding, address);
                adminClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                adminClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                resourceClient = new SvcResource.ResourceClient(binding, address);
                resourceClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                resourceClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                queueClient = new SvcQueueSystem.QueueSystemClient(binding, address);
                queueClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                queueClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;
            }
        }
    }
}
