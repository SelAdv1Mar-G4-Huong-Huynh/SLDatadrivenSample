using OpenQA.Selenium;
using System.Net;

namespace SeleniumAdvProject.Common
{
    public class Constants
    {
        //public static IWebDriver WebDriver;
        public const string Repository = "SampleRepository";
        public const string TestRepository = "TestRepository";
        public const string UserName = "administrator";
        public const string Password = "";
        public const int WaitTimeoutShortSeconds = 5;
        public const int lenghtRandomString = 10;
        public const string UserName1 = "test1";
        public const string SpecialPassword = "!@#$%^&*()";
        public const string SpecialUserName = "@()";
        public const string UpperCasePassword = "TEST";
        public const string UpperCaseUserName = "ADMIN";
        public const string LowerCaseUser = "admin";
        public const string UserName2 = "vutran";
        public const string Password2 = "VUTRAN";

        //URL
        //public const string LoginPageUrl = "http://192.168.1.104/TADashboard/login.jsp";
        public const string LoginPageUrl = "http://192.168.0.103:80/TADashboard/login.jsp";
        //public const string LoginPageUrl = "http://groupba.dyndns.org:54000/TADashboard/login.jsp";
        public const string MainPageUrl = "/TADashboard/";

        public static readonly string[] presetDataProfile = {"Action Implementation By Status",
                                     "Test Case Execution",
                                     "Test Case Execution Failure Trend",
                                     "Test Case Execution History",
                                     "Test Case Execution Results",
                                     "Test Case Execution Trend",
                                     "Test Module Execution",
                                     "Test Module Execution Failure Trend",
                                     "Test Module Execution History",
                                     "Test Module Execution Results",
                                     "Test Module Execution Results Report",
                                     "Test Module Execution Trend",
                                     "Test Module Implementation By Priority",
                                     "Test Module Implementation By Status",
                                     "Test Module Status per Assigned Users",
                                     "Test Objective Execution"};

    }

}
