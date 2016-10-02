using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Div: BaseControl
    {
        public Div() { }
        public Div(string xPath) : base(By.XPath(xPath)) { }
        public Div(IWebElement element) : base(element) { }
        public Div(By by) : base(by) { }
        public Div(IWebDriver webDriver, By by) : base(webDriver, by) { }
       
    }
}
