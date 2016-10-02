using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Label:BaseControl
    {  
        public Label() { }
        public Label(string xPath) : base(By.XPath(xPath)) { }
        public Label(IWebElement element) : base(element) { }
        public Label(By by) : base(by) { }
        public Label(IWebDriver webDriver, By by) : base(webDriver, by) { }
    }
}
