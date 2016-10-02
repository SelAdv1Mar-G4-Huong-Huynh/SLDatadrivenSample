using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class TextBox : BaseControl
    {
        public TextBox() { }
        public TextBox(string xPath) : base(By.XPath(xPath)) { }

        public TextBox(By by) : base(by) { }

        public TextBox(IWebElement element) : base(element) { }

        /// <summary>
        /// Returns whether textbox is ReadOnly.
        /// </summary>
        /// <returns></returns>
        public bool ReadOnly()
        {
            return bool.Parse(element.GetAttribute("readonly"));
        }

        /// <summary>
        /// Enters the text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// Author: Tu Nguyen
        public void EnterText(string text)
        {
            if (text == null)
            {
                return;
            }
            element.Click();
            element.SendKeys(text);
            string temp = element.Text;
            if (temp == null)
            {
                element.Click();
                element.SendKeys(text);
            }
        }

    }
}
