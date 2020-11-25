using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;

namespace FortyFingers.AddToCalendarHandler
{
    /// <summary>
    /// Summary description for AddToCalendar
    /// </summary>
    public class AddToCalendar : IHttpHandler
    {

        // /DesktopModules/40Fingers/AddToCalendarHandler/AddToCalendar.ashx?type=ical&subject=testsubject&description=testdescription&start=20201201T090000&duration=120&location=online

        public void ProcessRequest(HttpContext context)
        {
            var req = context.Request;
            var rsp = context.Response;
            var calendarType = req.QueryString["type"];

            var uniqueId = req.QueryString["uid"];
            if (string.IsNullOrEmpty(uniqueId)) uniqueId = Guid.NewGuid().ToString().Replace("-", "");
            var description = req.QueryString["description"];
            var location = req.QueryString["location"];
            var subject = req.QueryString["subject"];
            var startDate = DateTime.Now;
            if (DateTime.TryParseExact(req.QueryString["start"] ?? "", "yyyyMMddTHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate))
            {
                // it's parsed
            }
            var endDate = startDate.AddHours(1);
            if (!string.IsNullOrEmpty(req.QueryString["end"]))
            {
                DateTime.TryParseExact(req.QueryString["end"], "yyyyMMddTHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate);
            }
            int duration = 0;
            int.TryParse(req.QueryString["duration"], out duration);
            if (duration > 0)
            {
                endDate = startDate.AddMinutes(duration);
            }
            switch (calendarType.ToLower())
            {
                case "google":
                    // https://www.google.com/calendar/event?action=TEMPLATE&text=Eredienst&dates=20201101T100000/20201101T111500&details=&location=wonen+via+de+livestream%2C+'
                    var url = $"{"https://"}www.google.com/calendar/event?action=TEMPLATE&text={subject}&dates={startDate:yyyyMMddTHHmmss}/{endDate:yyyyMMddTHHmmss}&details={description}&location={location}";
                    rsp.Redirect(url);
                    break;
                case "icloud":
                case "outlook":
                case "ical":
                    var icsString = CreateIcs(subject,description, location, startDate, endDate, uniqueId);
                    var filename = "addToCalendar.ics";
                    rsp.Clear();
                    rsp.ClearHeaders();
                    rsp.ClearContent();
                    rsp.ContentType = "text/calendar";
                    rsp.AppendHeader("content-length", icsString.Length.ToString());
                    rsp.AppendHeader("Content-Disposition",
                        $"attachment; filename={filename}");
                    rsp.BinaryWrite(Encoding.ASCII.GetBytes(icsString));
                    rsp.Flush();
                    rsp.End();
                    break;
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

        //https://en.wikipedia.org/wiki/ICalendar

        private string CreateIcs(string subject, string description, string location, DateTime startDate, DateTime endDate, string uniqueId)
        {
            var icsString = $@"BEGIN:VCALENDAR
PRODID:-//Microsoft Corporation//Outlook 12.0 MIMEDIR//EN
VERSION:2.0
METHOD:PUBLISH
X-MS-OLK-FORCEINSPECTOROPEN:TRUE
BEGIN:VEVENT
CLASS:PUBLIC
DESCRIPTION:{subject}
DTEND:{endDate:yyyyMMddTHHmmss}
DTSTART:{startDate:yyyyMMddTHHmmss}
PRIORITY:5
SEQUENCE:0
SUMMARY;LANGUAGE=en-us:{description}
TRANSP:OPAQUE
UID:{uniqueId}
LOCATION:{location}
X-MICROSOFT-CDO-BUSYSTATUS:BUSY
X-MICROSOFT-CDO-IMPORTANCE:1
X-MICROSOFT-DISALLOW-COUNTER:FALSE
X-MS-OLK-ALLOWEXTERNCHECK:TRUE
X-MS-OLK-AUTOFILLLOCATION:FALSE
X-MS-OLK-CONFTYPE:0
END:VEVENT
END:VCALENDAR";

            return icsString;
        }

    }
}