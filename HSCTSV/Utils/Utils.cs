using DevExpress.CodeParser;
using DevExpress.XtraBars.Alerter;
using System;
using System.Net.Sockets;

namespace HSCTSV.Utils
{
    class Utils
    {
        public static string urlAPI = "http://202.191.56.101/HSCTSV";

        public static DateTime? convertTimeAPI(string timeText)
        {
            DateTime? date = null;
            try
            {
                date = DateTime.ParseExact(timeText,
                    "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture);
            }
            catch { }
            return date;
        }
        


    }
}
