using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Td: BaseControl
    {
        public Td() { }
        public Td(string xPath) : base(By.XPath(xPath)) { }
        public Td(IWebElement element) : base(element) { }
        public Td(By by) : base(by) { }
       
    }
}
