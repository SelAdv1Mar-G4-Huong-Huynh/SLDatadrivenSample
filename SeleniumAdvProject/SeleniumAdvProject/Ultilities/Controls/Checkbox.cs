using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Checkbox : BaseControl
    {
        public Checkbox() { }
        public Checkbox(string xPath) : base(By.XPath(xPath)) { }
        public Checkbox(IWebElement element) : base(element) { }
        public Checkbox(By by) : base(by) { }
                
        public void Check()
        {
            if (Selected == false)
            {
                Click();
            }
        }
                
        public void Uncheck()
        {
            if (Selected == true)
            {
                Click();
            }
        }
    }
}
