using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using SeleniumAdvProject.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SeleniumAdvProject.Ultilities;

namespace SeleniumAdvProject.PageObjects
{
    public class DashboardPage : BasePage
    {
        #region Locators

        static readonly By _lblUsername = By.XPath("//a[@href='#Welcome']");
        static readonly By _lblRepository = By.XPath("//a[@href='#Repository']");
        static readonly By _lblCurrentRepository = By.XPath("//a[@href='#Repository']/span");
        static readonly By _lnkLogout = By.XPath("//a[@href='logout.do']");
        static readonly By _lblSetting = By.XPath(".//li[@class='mn-setting']/a");

        #endregion

        #region Elements

        public IWebElement LblUsername
        {
            get { return Constant.WebDriver.FindElement(_lblUsername); }
        }
        public IWebElement LblRepository
        {
            get { return Constant.WebDriver.FindElement(_lblRepository); }
        }
        public IWebElement LblCurrentRepository
        {
            get { return Constant.WebDriver.FindElement(_lblCurrentRepository); }
        }
        public IWebElement LnkLogout
        {
            get { return Constant.WebDriver.FindElement(_lnkLogout); }
        }

        public IWebElement LblSetting
        {
            get { return Constant.WebDriver.FindElement(_lblSetting); }
        }
        #endregion

        #region Methods
        public string GetUserNameText()
        {
            return LblUsername.Text;
        }

        public string GetCurrentRepositoryText()
        {
            return LblCurrentRepository.Text;
        }

        public LoginPage Logout()
        {
            IWebElementExtensions.MoveTo(LblUsername, Constant.WebDriver);
            IWebElementExtensions.MoveTo(LnkLogout, Constant.WebDriver).Click();
            return new LoginPage();
        }

        public DashboardPage SelectRepository(String repositoryName)
        {
            Actions mouseAction = new Actions(Constant.WebDriver);
            mouseAction.MoveToElement(LblRepository).Perform();
            mouseAction.MoveToElement(Constant.WebDriver.FindElement(By.XPath("//ul[@id='ulListRepositories']//a[.='" + repositoryName + "']"))).Click().Perform();
            return this;
        }

        public AddnEditPage GoToAddNewPage()
        {
            Actions mouseAction = new Actions(Constant.WebDriver);
            mouseAction.MoveToElement(LblSetting).Perform();
            mouseAction.MoveToElement(Constant.WebDriver.FindElement(By.XPath("//a[@class='add']"))).Click().Perform();
            WaitForControlExists(By.XPath(".//*[@id='parent']"), 2);
            return new AddnEditPage();
        }

        public DashboardPage OpenSetting()
        {
            Actions mouseAction = new Actions(Constant.WebDriver);
            mouseAction.MoveToElement(LblSetting).Perform();
            return this;
        }

        public Boolean IsSettingExist()
        {
            bool _isExist = false;
            _isExist = LblSetting.Selected;
            return _isExist;
        }


        #endregion
    }
}
