using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using SeleniumAdvProject.Common;
using SeleniumAdvProject.DataObjects;
using OpenQA.Selenium.Support.UI;
using SeleniumAdvProject.Ultilities.Controls;
using System.Threading;

namespace SeleniumAdvProject.PageObjects
{
    public class SortFieldsPage : DataProfileBasePage
    {
        #region Locators
        static readonly By _cbbField = By.XPath("//select[@id='cbbFields']");
        static readonly By _btnAddLevel = By.XPath("//input[@id='btnAddSortField']");
        static readonly By _cbbDateField = By.XPath("//select[@name='cbbSortInterval']");
        #endregion

        #region Elements
        public ComboBox CbbField
        {
            get { return new ComboBox(FindElement(_cbbField)); }
        }
        public ComboBox CbbDateField
        {
            get { return new ComboBox(FindElement(_cbbDateField)); }
        }
        public Button BtnAddLevel
        {
            get { return new Button(FindElement(_btnAddLevel)); }
        }
        #endregion

        #region Methods
        public SortFieldsPage() { }
        public SortFieldsPage(IWebDriver webDriver) : base(webDriver) { }

        public bool IsItemTypeListed(string[] expectedItemType)
        {
            IList<string> itemTypeList = CbbField.OptionStrings;
            foreach (string item in expectedItemType)
            {
                if (!itemTypeList.Contains(item))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Determines whether [is field level exist] [the specified field name].
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public bool IsFieldLevelExist(string fieldName)
        {
            Span fieldLevel = new Span(FindElement(By.XPath(string.Format(".//table[@id='profilesettings']//span[.='{0}']", fieldName))));
            if (fieldLevel == null)
                return false;
            return true;
        }

        /// <summary>
        /// Adds the level.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public SortFieldsPage AddLevel(string fieldName)
        {
            CbbField.SelectByTextFromGroup(fieldName);
            BtnAddLevel.Click();
            return this;

        }

        /// <summary>
        /// Adds the level with expected error.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string AddLevelWithExpectedError(string fieldName)
        {
            CbbField.SelectByTextFromGroup(fieldName);
            BtnAddLevel.Click();
            return this.GetDialogText();
        }

        /// <summary>
        /// Removes the field level.
        /// </summary>
        /// <param name="fieldName">Name of the field.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public SortFieldsPage RemoveFieldLevel(string fieldName)
        {
            Button removeButton = new Button(FindElement(By.XPath(string.Format("//table[@id='profilesettings']//td/span[.='{0}']/../../td/button[@title='Remove']", fieldName))));
            removeButton.Click();
            return this;
        }
        #endregion
    }
}
