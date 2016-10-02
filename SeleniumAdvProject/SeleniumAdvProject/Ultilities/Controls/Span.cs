using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Span : BaseControl
    {
        public Span() { }
        public Span(string xPath) : base(By.XPath(xPath)) { }
        public Span(IWebElement element) : base(element) { }
        public Span(By by) : base(by) { }
        public Span(IWebDriver webDriver, By by) : base(webDriver, by) { }
    }
}
