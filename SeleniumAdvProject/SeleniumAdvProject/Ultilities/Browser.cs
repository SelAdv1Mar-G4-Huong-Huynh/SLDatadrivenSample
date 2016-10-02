using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumAdvProject.Common;
using SeleniumAdvProject.Ultilities.Controls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities
{
    public class Browser
    {
        protected IWebDriver _webDriver;
        protected IWebDriver webDriver
        {
            get { return this._webDriver; }
            set { this._webDriver = value; }
        }

        #region Constructor
        public Browser() { }
        public Browser(IWebDriver webDriver)
        {
            this._webDriver = webDriver;
        }
        #endregion

        #region Method

        public IWebElement FindElement(By by, int waitingTime = 60)
        {

            IWebElement iElement = null;
            //int waitSecond = 0;

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            if (stopwatch.Elapsed.Seconds < waitingTime)
            {
                try
                {
                    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(waitingTime));
                    wait.Until(ExpectedConditions.ElementIsVisible(by));
                    iElement = _webDriver.FindElement(by);                  
                }
                catch (StaleElementReferenceException)
                {
                    waitingTime = waitingTime - stopwatch.Elapsed.Seconds;
                    FindElement(by, waitingTime);
                }
                catch (NullReferenceException)
                {
                    waitingTime = waitingTime - stopwatch.Elapsed.Seconds;
                    FindElement(by, waitingTime);
                }
                catch (WebDriverException)
                {
                    waitingTime = waitingTime - stopwatch.Elapsed.Seconds;
                    FindElement(by, waitingTime);
                }
            }           
            stopwatch.Stop();

            return iElement;
        }

        #endregion
    }
}
