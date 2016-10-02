using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Link : BaseControl
    {
        public Link() { }        
        public Link(string xPath) : base(By.XPath(xPath)) { }
        public Link(IWebElement element) : base(element) { }
        public Link(By by) : base(by) { }
        public Link(IWebDriver webDriver, By by) : base(webDriver, by) { }
    }
}
