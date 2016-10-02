using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using SeleniumAdvProject.Common;
using OpenQA.Selenium.Support.UI;
using System.Net;
using SeleniumAdvProject.Ultilities.Controls;

namespace SeleniumAdvProject.PageObjects
{
    public class LoginPage : BasePage
    {
        #region Locators
        static readonly By _ddlRepository = By.XPath("//select[@name='repository']");
        static readonly By _txtUsername = By.XPath("//input[@id='username']");
        static readonly By _txtPassword = By.XPath("//input[@id='password']");
        static readonly By _btnLogin = By.XPath("//div[@class='btn-login']");
        #endregion

        #region Elements
        public ComboBox DdlRepsitory
        {
            get { return new ComboBox(FindElement(_ddlRepository)); }
        }
        public TextBox TxtUsername
        {
            get { return new TextBox(FindElement(_txtUsername)); }
        }
        public TextBox TxtPassword
        {
            get { return new TextBox(FindElement(_txtPassword)); }
        }
        public Button BtnLogin
        {
            get { return new Button(FindElement(_btnLogin)); }
        }
        #endregion

        #region Methods

        public LoginPage() { }

        public LoginPage(IWebDriver webDriver) : base(webDriver) { }

        /// <summary>
        /// Opens Login Page
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public LoginPage Open()
        {
            _webDriver.Navigate().GoToUrl(Constants.LoginPageUrl);
            return this;
        }

        /// <summary>
        /// Logins the dashboard web app.
        /// </summary>
        /// <param name="repository">The repository name</param>
        /// <param name="username">The username</param>
        /// <param name="password">The password</param>
        /// <returns>The MainPage object</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public MainPage Login(string repository, string username, string password)
        {
            DdlRepsitory.SelectByText(repository);
            TxtUsername.SendKeys(username);
            TxtPassword.SendKeys(password);
            BtnLogin.Click();           
            return new MainPage(_webDriver);
        }

        /// <summary>
        /// Logins dashboard web app with expected error.
        /// </summary>
        /// <param name="repository">The repository.</param>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        /// <returns> The error message</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public string LoginWithExpectedError(string repository, string username, string password)
        {
            DdlRepsitory.SelectByText(repository);
            TxtUsername.SendKeys(username);
            TxtPassword.SendKeys(password);
            BtnLogin.Click();
            return this.GetDialogText();
        }

        /// <summary>
        /// Logins the with out account.
        /// </summary>
        /// <param name="repository">The repository.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string LoginWithOutAccount(string repository)
        {
            DdlRepsitory.SelectByText(repository);
            BtnLogin.Click();
            return this.GetDialogText();
        }

        /// <summary>
        /// Gets the web driver.
        /// </summary>
        /// <returns></returns>
        public IWebDriver GetWebDriver()
        {
            return _webDriver;
        }
        #endregion

        /// <summary>
        /// Verify Login page is displayed
        /// </summary>
        /// <returns>True/False</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2015</date>
        public bool Displayed()
        {
            Button BtnLogin = new Button(_webDriver, _btnLogin);
            return BtnLogin.isDisplayed();
        }

    }
}
