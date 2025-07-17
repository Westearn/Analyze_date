using System;
using System.Collections.Generic;
using System.Threading;


namespace ANData
{
    static class variables
    {
        public static string path = AppDomain.CurrentDomain.BaseDirectory;
        public static string[] directory { get; set; }
        public static List<DateTime> dateTimeIn { get; set; } = new List<DateTime>();
        public static List<DateTime> dateTimeOut { get; set; } = new List<DateTime>();
        public static List<DateTime> dateTimepd { get; set; } = new List<DateTime>();
        public static List<string> number_kp { get; set; }
        public static CancellationTokenSource source;
        public static CancellationToken token;
    }
}
