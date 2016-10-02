using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Button : BaseControl
    {
        public Button() { }
        public Button(string xPath) : base(By.XPath(xPath)) { }
        public Button(IWebElement element) : base(element) { }
        public Button(By by) : base(by) { }
        public Button(IWebDriver webDriver, By by) : base(webDriver, by) { }
    }
}
