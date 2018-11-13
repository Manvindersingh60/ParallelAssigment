using ExcelDna.Integration;
using System;
using System.Net.Http;
using System.Windows.Forms;

namespace ParallelAssigment
{
    /// <summary>
    /// Contains BL methods
    /// </summary>
    static class Utility
    {
        #region Private fields

        private static int _requestCount;
        private static readonly object _showObject = new object();
        private static readonly object _hideObject = new object();
        private static DateTime _stopTime;
        private static readonly string _url = System.Configuration.ConfigurationSettings.AppSettings["url"];

        #endregion

        public static ProcessingForm Form { get; set; }
        public static int SuccessfulHits { get; set; }

        /// <summary>
        /// Send the web api request and generates response
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public static object SendRequest(int num)
        {
            //Don't serve any request if previous stop was less than 1 sec ago(time can be changes according to requirement)
            if(DateTime.Now.Subtract(_stopTime).Seconds < 1)
            {
                return ExcelError.ExcelErrorNull;
            }

            lock (_showObject)
            {
                _requestCount++;
                Form.BeginInvoke(new MethodInvoker(Form.Show));
            }

            try
            {
                using (var client = new HttpClient())
                {
                    var response = client.GetAsync(string.Format(_url, num)).Result;

                    if(response.IsSuccessStatusCode)
                    {
                        SuccessfulHits++;
                    }
                }
            }
            catch (Exception)
            {
            }

            lock (_hideObject)
            {
                if(_requestCount > 0)
                {
                    _requestCount--;
                }

                if (_requestCount == 0)
                {
                    Form.BeginInvoke(new MethodInvoker(Form.Hide));
                }
            }

            return num + 2;
        }

        /// <summary>
        /// Stops all the pending requests
        /// </summary>
        public static void StopRequest()
        {
            Form.BeginInvoke(new MethodInvoker(Form.Hide));
            _requestCount = 0;
            _stopTime = DateTime.UtcNow;
        }
    }

    class Model
    {
        public int Num { get; set; }
        public string Cell { get; set; }
    }
}
