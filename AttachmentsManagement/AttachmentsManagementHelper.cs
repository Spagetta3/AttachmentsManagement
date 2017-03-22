using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentsManagementAgent
{
    class AttachmentsManagementHelper
    {
        private const string s_EventLogName = "Application";

        public static void WriteLog(string message, EventLogEntryType entryType,
                                int eventID, string proccessName)
        {

            try
            {
                EventLog evtLog = new EventLog();
                evtLog.Log = s_EventLogName;
                evtLog.Source = proccessName;
                if (!EventLog.SourceExists(evtLog.Source))
                {
                    EventLog.CreateEventSource(evtLog.Source, evtLog.Log);
                }
                evtLog.WriteEntry(message, entryType, eventID);
            }
            catch (ArgumentException)
            {
            }
            catch (InvalidOperationException)
            {
            }
        }
    }
}
