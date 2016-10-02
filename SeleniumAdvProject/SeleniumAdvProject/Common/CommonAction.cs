using System;
using System.Threading;

namespace SeleniumAdvProject.Common
{
    public static class CommonAction
    {
        /// <summary>
        /// Genrate the random string
        /// </summary>
        /// <param name="length">The length of string</param>
        /// <returns>String</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public static string GenrateRandomString(int length = Constants.lenghtRandomString)
        {
            return Guid.NewGuid().ToString().Substring(0, length);
        }

        /// <summary>
        /// Genrate the string by Date
        /// </summary>
        /// <param name="day">The number of days from today</param>
        /// <returns>String</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public static string RandomDay(int day)
        {
            return DateTime.Now.AddDays(day).ToString("M/dd/yyy");
        }

        /// <summary>
        /// Encode the space to \u00a0
        /// </summary>
        /// <param name="original">The original string</param>
        /// <returns>String</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2015</date>
        public static string EncodeSpace(string original)
        {
            return original.Replace(' ', '\u00a0');
        }

        /// <summary>
        /// Generate the page name
        /// </summary>
        /// <returns>String</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2016</date>
        public static string GeneratePageName()
        {
            return string.Format("Page {0}", GenrateRandomString());
        }

        /// <summary>
        /// Generate the Panel name
        /// </summary>
        /// <returns>String</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2016</date>
        public static string GeneratePanelName()
        {
            return string.Format("Panel {0}", GenrateRandomString());
        }

        /// <summary>
        /// Generate the Data Profile name
        /// </summary>
        /// <returns>String</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2016</date>
        public static string GenerateDataProfileName()
        {
            return string.Format("DataProfile {0}", GenrateRandomString());
        }
    }
}
