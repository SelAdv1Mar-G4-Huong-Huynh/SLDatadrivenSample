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

namespace SeleniumAdvProject.PageObjects
{
    public class Popup : BasePage
    {        
        #region Locators       
        static readonly By _btnOk = By.XPath("//input[@id='OK']");
        static readonly By _btnCancel = By.XPath("//input[@id='Cancel']");
        #endregion

        #region Elements       
        public Button BtnOk
        {
            get { return new Button(FindElement(_btnOk)); }
        }
        public Button BtnCancel
        {
            get { return new Button(FindElement(_btnCancel)); }
        }
        #endregion

        #region Methods
        public Popup() { }
        public Popup(IWebDriver webDriver) : base(webDriver) { }

        #endregion


    }
}
