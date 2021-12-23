using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Diagnostics;

namespace ExCoreService
{

    class EventLogUtil
    {
        public static EventLog EventLogRef;
        public static int TraceLevel = int.Parse(ConfigurationManager.AppSettings["TraceLevel"]);
        public static string LastLog;
        public static string LastLogError;

        private static string sEventSource = typeof(Program).Assembly.GetName().Name + "Source";
        private static string sEventLog = typeof(Program).Assembly.GetName().Name + "Log";
        public static void InitEventLog()
        {
            EventLogUtil.EventLogRef = new EventLog();

            foreach(EventLog oEventLog in EventLog.GetEventLogs())
            {
                if(sEventLog.Substring(0,8).ToUpper() == oEventLog.Log.Substring(0, 8).ToUpper())
                {
                    sEventLog = oEventLog.Log;
                    break;
                }
            }
            if (!EventLog.SourceExists(sEventSource))
            {
                EventLog.CreateEventSource(sEventSource, sEventLog);
            }
        }
        public static void LogEvent(string EventMessage, EventLogEntryType EventType = EventLogEntryType.Information, int _EventId = 0)
        {
            if (EventLogUtil.EventLogRef is null) InitEventLog();
            EventLogUtil.EventLogRef.Source = sEventSource;
            EventLogUtil.EventLogRef.Log = sEventLog;
            if (_EventId <= TraceLevel)
            {
                EventLogUtil.EventLogRef.WriteEntry(EventMessage, EventType, _EventId);
                LastLog = EventMessage;
            }
        }
        public static void LogErrorEvent(string Location, Exception Ex, int _Event_Id = 0)
        {
            if(EventLogUtil.EventLogRef is null) InitEventLog();
            EventLogUtil.EventLogRef.Source = sEventSource;
            EventLogUtil.EventLogRef.Log = sEventLog;
            LastLogError = "Error Location: " + Location + "\r\n" + "Error Source: " + Ex.Source + "\r\n" + "Error Message: " + Ex.Message;
            EventLogUtil.EventLogRef.WriteEntry(LastLogError, EventLogEntryType.Error, _Event_Id);
        }
    }
}
