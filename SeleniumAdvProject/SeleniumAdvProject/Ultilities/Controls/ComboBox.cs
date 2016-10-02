using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class ComboBox : BaseControl
    {
        private SelectElement selectElement;
        public ComboBox() { }

        public ComboBox(string xPath) : base(By.XPath(xPath)) { }

        public ComboBox(By by) : base(by) { }

        public ComboBox(IWebElement element) : base(element) { }

        /// <summary>
        /// get elect the control
        /// </summary>
        /// <returns></returns>
        private void GetSelectControl()
        {
            //LoadControl();
            selectElement = new SelectElement(element);
        }
        public bool isEnabled
        {
            get
            {
                //LoadControl();
                return element.Selected;
            }
        }

        /// <summary>
        /// Selects the by text.
        /// </summary>
        /// <param name="text">The text.</param>
        public void SelectByText(string text)
        {
            GetSelectControl();
                if (text != null)
                    selectElement.SelectByText(text);

        }
        /// <summary>
        /// Selects the index of the by.
        /// </summary>
        /// <param name="index">The index.</param>
        public void SelectByIndex(int index)
        {
            //GetSelectControl();
            selectElement = new SelectElement(element);
            selectElement.SelectByIndex(index);
        }

        /// <summary>
        /// Gets the selected text.
        /// </summary>
        /// <returns></returns>
        public string GetSelectedText()
        {
            GetSelectControl();
            return selectElement.SelectedOption.Text;
        }

        /// <summary>
        /// Gets the option strings.
        /// </summary>
        /// <value>
        /// The option strings.
        /// </value>
        public IList<string> OptionStrings
        {
            get
            {
                GetSelectControl();
                IList<IWebElement> options = selectElement.Options;
                return options.Select(option => option.Text).ToList();
            }
        }

        /// <summary>
        /// Selects the by text from group.
        /// </summary>
        /// <param name="text">The text.</param>
        public void SelectByTextFromGroup(string text)
        {
            IList<IWebElement> options = element.FindElements(By.TagName("option"));
            foreach (IWebElement option in options)
            {
                if (option.Text.Trim().Equals(text))
                {
                    option.Click();
                    return;
                }
            }
        }

        /// <summary>
        /// Gets the index of the text by.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetTextByIndex(int index)
        {
            IList<string> text = OptionStrings;
            return text[index];
        }
    }
}
