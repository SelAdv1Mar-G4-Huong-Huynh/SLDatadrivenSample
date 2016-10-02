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
    public class DisplayFieldsPage : DataProfileBasePage
    {
        #region Locators
        static readonly By _lnkCheckAll = By.XPath("//table[@id='profilesettings']//a[.='Check All']");
        static readonly By _lnkUnCheckAll = By.XPath("//table[@id='profilesettings']//a[.='Uncheck All']");
        #endregion

        #region Elements
        public Link LnkCheckAll
        {
            get { return new Link(FindElement(_lnkCheckAll)); }
        }
        public Link LnkUnCheckAll
        {
            get { return new Link(FindElement(_lnkUnCheckAll)); }
        }
        #endregion


        #region Methods

        public DisplayFieldsPage() : base() { }

        public DisplayFieldsPage(IWebDriver webDriver) : base(webDriver) { }

        public bool IsCheckBoxFieldDisplayed(string checkboxName)
        {
            Checkbox chkField = new Checkbox(FindElement(By.XPath(string.Format("//table[@id='profilesettings']//label[text()=' {0}']/input[@type='checkbox']", checkboxName))));
            if (chkField == null)
                return false;
            return true;
        }

        public bool IsCheckBoxChecked(string checkboxName)
        {
            Checkbox chkField = new Checkbox(FindElement(By.XPath(string.Format("//table[@id='profilesettings']//label[text()=' {0}']/input[@type='checkbox']", checkboxName))));
            if (GetCheckBoxStatus(chkField))
                return true;
            return false;
        } 

        public bool IsFieldDisplayed(string fieldName)
        {
            Label lblName = new Label(FindElement(By.XPath(string.Format("//table[@id='profilesettings']//label[text()=' {0}']", fieldName))));
            if (lblName == null)
                return false;
            return true;
        }
        public DisplayFieldsPage ClickCheckAll()
        {
            LnkCheckAll.Click();
            return this;
        }

        public DisplayFieldsPage ClickUnCheckAll()
        {
            LnkUnCheckAll.Click();
            return this;
        }
        
        #endregion
    }
}
