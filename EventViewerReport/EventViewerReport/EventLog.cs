using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics.Eventing.Reader;

namespace EventViewerReport
{
    class EventLogHelper
    {
        private readonly string FilterSystem;
        private readonly string FilterApplication;

        public EventLogHelper()
        {
            Utility util = new Utility();
            string propertyFileName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".properties";
            string TimeInterval = util.GetPropertyValue("TimeInterval", propertyFileName);
            int MilliSecondsInterval = Convert.ToInt32(TimeInterval) * 60 * 1000;
            FilterSystem =
                 "<QueryList>" +
                 "  <Query Id=\"0\" Path=\"System\">" +
                 "    <Select Path=\"System\">*[System[(Level=1  or Level=2) and TimeCreated[timediff(@SystemTime) &lt;= " + MilliSecondsInterval + "]]]</Select>" +
                 "  </Query>" +
                 "</QueryList>";

            FilterApplication =
                "<QueryList>" +
                "  <Query Id=\"0\" Path=\"Application\">" +
                "    <Select Path=\"Application\">*[System[(Level=1  or Level=2) and TimeCreated[timediff(@SystemTime) &lt;= " + MilliSecondsInterval + "]]]</Select>" +
                "  </Query>" +
                "</QueryList>";
        }

        public string GetRecentEvents(string remoteIpAddress)
        {
            string EventLogFolder = "System";
            string returnValue = "";

            var query = BuildQuery(EventLogFolder, remoteIpAddress, FilterSystem);
            returnValue = ReadLogs(new EventLogReader(query));
            if (returnValue == "")
            {
                EventLogFolder = "Application";
                query = BuildQuery(EventLogFolder, remoteIpAddress, FilterApplication);
                returnValue = ReadLogs(new EventLogReader(query));
            }
            if (returnValue == "")
            {
                returnValue = "        No Events Error Found in this Machine.";
            }
            return returnValue;

        }


        /// <summary>
        /// Builds an EventLogQuery for the given pcname and filter. This user needs to be in the Event Log Readers security group. 
        /// </summary>
        private static EventLogQuery BuildQuery(string eventLogFolder, string pcName, string filter)
        {
            var session = new EventLogSession();
            bool isSessionGood = false;

            if (!isSessionGood)
            {
                try
                {
                    using (var pw = GetPassword())
                    {
                        session = new EventLogSession(
                        pcName,
                        pcName,
                        "43Administrator",
                        pw,
                        SessionAuthentication.Default);
                    }
                    isSessionGood = true;
                }
                catch (UnauthorizedAccessException e)
                {
                    //  requires elevation of privilege
                    Console.WriteLine("You do not have the correct permissions, Try running with administrator privileges. " + e.ToString());
                }
            }

            return new EventLogQuery(eventLogFolder, PathType.LogName, filter)
            { Session = session };
        }


        /// <summary>
        /// Read the given EventLogReader and return the amount of events that match the IDs we are looking for
        /// </summary>
        public string ReadLogs(EventLogReader logReader)
        {
            string returnValue = "";

            try
            {
                for (EventRecord eventdetail = logReader.ReadEvent(); eventdetail != null; eventdetail = logReader.ReadEvent())
                {
                    returnValue += "\t" + "<li>" + eventdetail.TimeCreated + " " + eventdetail.FormatDescription() + "</li>" + "\r\n";
                }
            }
            catch (Exception e)
            {
                returnValue = "\t" + "<li>" + "EXCEPTION: " + e.ToString() + "</li>" + "\r\n";
            }

            return returnValue;
        }


        /// <summary>
        /// Return the password stored in an encrypted bytearray. 
        /// </summary>
        private static System.Security.SecureString GetPassword()
        {
            Utility util = new Utility();

            string pswd = util.Decrypt("2RleqdHjMLseKHFzRNq4Bg****");
            var ss = new System.Security.SecureString();

            foreach (char c in pswd)
            {
                ss.AppendChar(c);
            }

            return ss;
        }
    }
}

