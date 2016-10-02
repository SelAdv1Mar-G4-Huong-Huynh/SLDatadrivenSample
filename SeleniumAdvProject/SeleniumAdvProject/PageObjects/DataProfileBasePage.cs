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
    public class DataProfileBasePage : BasePage
    {
        #region Locators
        protected static readonly By _btnNext = By.XPath("//input[@value='Next' and @type='button']");
        protected static readonly By _btnFinish = By.XPath("//input[@value='Finish' and @type='button']");
        protected static readonly By _btnCancel = By.XPath("//input[@value='Cancel' and @type='button']");
        protected static readonly By _btnBack = By.XPath("//input[@value='Back' and @type='button']");
        protected static readonly By _lblSortField = By.XPath("//ul[@id='wstep']/li[text()='Sort Fields']");
        protected static readonly By _lblGeneralSetting = By.XPath("//ul[@id='wstep']/li[text()='General Settings']");
        protected static readonly By _lblDisplayField = By.XPath("//ul[@id='wstep']/li[text()='Display Fields']");
        protected static readonly By _lblFilterFields = By.XPath("//ul[@id='wstep']/li[text()='Filter Fields']");
        protected static readonly By _lblStatisticFields = By.XPath("//ul[@id='wstep']/li[text()='Statistic Fields']");
        protected static readonly By _lblDisplaySubFiled = By.XPath("//ul[@id='wstep']/li[text()='Display Sub-Fields']");
        protected static readonly By _lblSortSubField = By.XPath("//ul[@id='wstep']/li[text()='Sort Sub-Fields']");
        protected static readonly By _lblFilterSubField = By.XPath("//ul[@id='wstep']/li[text()='Filter Sub-Fields']");
        protected static readonly By _lblStatisticSubField = By.XPath("//ul[@id='wstep']/li[text()='Statistic Sub-Fields']");
        protected static readonly By _lblProfileHeader = By.XPath("//td[@class='profilesettingheader']");
        
        #endregion

        #region Elements
        public Button BtnNext
        {
            get { return new Button(FindElement(_btnNext)); }
        }
        public Label LblProfileHeader
        {
            get { return new Label(FindElement(_lblProfileHeader)); }
        }
        public Label LblSortField
        {
            get { return new Label(FindElement(_lblSortField)); }
        }
        public Label LblGeneralSetting
        {
            get { return new Label(FindElement(_lblGeneralSetting)); }
        }
       
        public Label LblDisplayField
        {
            get { return new Label(FindElement(_lblDisplayField)); }
        }        
        public Label LblFilterFields
        {
            get { return new Label(FindElement(_lblFilterFields)); }
        }
        
        public Label LblStatisticFields
        {
            get { return new Label(FindElement(_lblStatisticFields)); }
        }
        
        public Label LblDisplaySubFiled
        {
            get { return new Label(FindElement(_lblDisplaySubFiled)); }
        }
        
        public Label LblSortSubField
        {
            get { return new Label(FindElement(_lblSortSubField)); }
        }
        
        public Label LblFilterSubField
        {
            get { return new Label(FindElement(_lblFilterSubField)); }
        }
        
        public Label LblStatisticSubField
        {
            get { return new Label(FindElement(_lblStatisticSubField)); }
        }
        public Button BtnFinish
        {
            get { return new Button(FindElement(_btnFinish)); }
        }
        public Button BtnCancel
        {
            get { return new Button(FindElement(_btnCancel)); }
        }
        public Button BtnBack
        {
            get { return new Button(FindElement(_btnBack)); }
        }
        #endregion

        #region Methods
        public DataProfileBasePage() : base() { }

        public DataProfileBasePage(IWebDriver webDriver) : base(webDriver) { }

        /// <summary>
        /// Click on the button
        /// </summary>
        /// <param name="buttonName">Name of the button</param>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        protected void ClickOnButton(string buttonName)
        {
            switch (buttonName)
            {
                case "Back":
                    BtnBack.Click();
                    break;
                case "Next":
                    BtnNext.Click();
                    break;
                case "Finish":
                    BtnFinish.Click();
                    break;
                case "Cancel":
                    BtnCancel.Click();
                    break;
            }
        }


        public GeneralSettingsPage GoToGeneralSettingsPage()
        {
            LblGeneralSetting.Click();
            return new GeneralSettingsPage(_webDriver);
        }

        public SortFieldsPage GoToSortFieldsPage()
        {
            Thread.Sleep(500);
            LblSortField.Click();
            return new SortFieldsPage(_webDriver);
        }
        #endregion

    }
}
