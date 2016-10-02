using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class RadioButton: BaseControl
    {
        public RadioButton() { }
        public RadioButton(string xPath) : base(By.XPath(xPath)) { }
        public RadioButton(IWebElement element) : base(element) { }
        public RadioButton(By by) : base(by) { }
       
    }
}
