using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumAdvProject.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SeleniumAdvProject.Ultilities.Controls;
using SeleniumAdvProject.Ultilities;

namespace SeleniumAdvProject.PageObjects
{
    public class BasePage:Browser
    {
        
        #region Locators
        protected static readonly By _lblUsername = By.XPath("//a[@href='#Welcome']");
        static readonly By _lblRepository = By.XPath("//a[@href='#Repository']");
        static readonly By _lblCurrentRepository = By.XPath("//a[@href='#Repository']/span");
        static readonly By _lnkLogout = By.XPath("//a[@href='logout.do']");
        protected static readonly By _lblGlobalSetting = By.XPath("//div[@id='main-menu']//li[@class='mn-setting']/a");
        static readonly By _lnkAddPage = By.XPath("//a[@class='add' and .='Add Page']");
        static readonly By _lnkEditPage = By.XPath("//a[@class='edit' and .='Edit']");
        static readonly By _lnkCreateProfile = By.XPath("//a[@class='add' and .='Create Profile']");
        static readonly By _lnkCreatePanel = By.XPath("//a[@class='add' and .='Create Panel']");
        static readonly By _btnChoosePanel = By.XPath("//a[@id='btnChoosepanel']");
        static readonly By _lnkAdminister = By.XPath("//a[.='Administer']");
        static readonly By _lnkPanel = By.XPath("//a[.='Panels']");
        static readonly By _lnkDataProfiles = By.XPath("//a[.='Data Profiles']");
        static readonly By _btnCreateNewPanel = By.XPath("//span[.='Create new panel']");
        static readonly By _lnkEditPanel = By.XPath("//li[@class='edit' and @title='Edit Panel']");

        #endregion

        #region Elements
        public Label LblRepository
        {
            get { return new Label(FindElement(_lblRepository));}
        }

        public Link LnkEditPanel
        {
            get { return new Link(FindElement(_lnkEditPanel)); }
        }
        public Link LnkDataProfiles
        {
            get { return new Link(FindElement(_lnkDataProfiles)); }
        }

        public Button BtnCreateNewPanel
        {
            get { return new Button(FindElement(_btnCreateNewPanel)); }
        }
        public Label LblCurrentRepository
        {
            get { return new Label(FindElement(_lblCurrentRepository)); }
        }
        public Label LblGlobalSetting
        {
            get { return new Label(FindElement(_lblGlobalSetting)); }
        }
        public Link LnkAddPage
        {
            get { return new Link(FindElement(_lnkAddPage)); }
        }
        public Link LnkCreatePanel
        {
            get { return new Link(FindElement(_lnkCreatePanel)); }
        }
        public Link LnkEditPage
        {
            get { return new Link(FindElement(_lnkEditPage)); }
        }

        public Button BtnChoosePanel
        {
            get { return new Button(FindElement(_btnChoosePanel)); }
        }
        public Link LnkLogout
        {
            get { return new Link(FindElement(_lnkLogout)); }
        }
        public Link LblUsername
        {
            get { return new Link(FindElement(_lblUsername)); }
        }
        public Link LnkAdminister
        {
            get { return new Link(FindElement(_lnkAdminister)); }
        }
        public Link LnkPanel
        {
            get { return new Link(FindElement(_lnkPanel)); }
        }
        #endregion

        #region Methods

        public BasePage() { }
        public BasePage(IWebDriver webDriver) : base(webDriver) { }
      
        #region Navigate Methods

        /// <summary>
        /// Go to the page link follow path of page
        /// </summary>
        /// <param name="pathOfPage">The path to go to the current page</param>
        /// <param name="isClicked">if set to <c>true</c> click on final link</param>
        /// <returns></returns>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        public void GoToLink(string pathOfPage, bool isClicked = false)
        {

            string[] arrNode = pathOfPage.Split('/');
            Link LnkName = new Link();
            foreach (string node in arrNode)
            {
                LnkName = new Link(FindElement(By.XPath(string.Format("//a[.='{0}']", CommonAction.EncodeSpace(node)))));
                LnkName.MouseOver();
            }
            if (isClicked)
            {
                LnkName.Click();
            }
        }

        public MainPage GoToPage(string path)
        {
            GoToLink(path, true);
            return new MainPage(_webDriver);
        }

        /// <summary>
        /// Opens the add new page by click on Global Setting>Add Page menu
        /// </summary>
        /// <returns> add new page popup </returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public AddNewPage OpenAddNewPage()
        {
            LblGlobalSetting.MouseOver();
            LnkAddPage.Click();
            return new AddNewPage(_webDriver);
        }

        /// <summary>
        /// Determines the combobox contains the exptected items
        /// </summary>
        /// <param name="comBoboxName">Name of the COM bobox.</param>
        /// <param name="listValues">The list values need to check</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>6/3/2016</date>
        public bool isComboboxContainsItems(ComboBox comBoboxName, string[] listValues)
        {
            bool flag = false;
            IList<String> values = comBoboxName.OptionStrings;
            foreach (string listValue in listValues)
            {
                flag = values.Contains(listValue);
                if (flag == false)
                    break;
            }
            return flag;
        }
        /// <summary>
        /// Opens the edit page.
        /// </summary>
        /// <param name="pageName">Name of the page.</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public AddNewPage OpenEditPage(string pageName)
        {
            GoToLink(pageName, true);
            LblGlobalSetting.MouseOver();
            LnkEditPage.Click();
            return new AddNewPage(_webDriver);
        }
        /// <summary>
        /// Opens the add new panel popup.
        /// </summary>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        public AddNewPanelPage OpenAddNewPanelPopup()
        {
            LblGlobalSetting.MouseOver();
            LnkCreatePanel.Click();
            return new AddNewPanelPage(_webDriver);
        }
        /// <summary>
        /// Opens the panels page.
        /// </summary>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        public PanelsPage OpenPanelsPage()
        {
            LnkAdminister.MouseOver();
            LnkPanel.Click();
            return new PanelsPage(_webDriver);
        }

        /// <summary>
        /// Go to the Data Profile page.
        /// </summary>
        /// <returns></returns>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        public DataProfilePage GoToDataProfilePage()
        {
            LnkAdminister.MouseOver();
            LnkDataProfiles.Click();
            return new DataProfilePage(_webDriver);
        }

        /// <summary>
        /// Go to the Data Profile page.
        /// </summary>
        /// <returns></returns>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        public AddNewPanelPage OpenAddNewPanelPage(string openFrom = "Global setting")
        {
            switch (openFrom)
            {
                case "Global setting":
                    {
                        LblGlobalSetting.MouseOver();
                        LnkCreatePanel.Click();
                        break;
                    }
                case "Choose panels":
                    {
                        BtnChoosePanel.Click();
                        BtnCreateNewPanel.Click();
                        break;
                    }
                case "Panels Page":
                    {
                        PanelsPage panelPage = OpenPanelsPage();
                        panelPage.LnkAddNew.Click();
                        break;
                    }
            }

            return new AddNewPanelPage(_webDriver);
        }

        /// <summary>
        /// Logouts this instance.
        /// </summary>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        public LoginPage Logout()
        {
            LblUsername.MouseOver();
            LnkLogout.Click();
            return new LoginPage(_webDriver);
        }


        /// <summary>
        /// Selects the repository on Main Page
        /// </summary>
        /// <param name="repositoryName">Name of the repository</param>
        /// <returns>The MainPage object</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2016</date>
        public MainPage SelectRepository(String repositoryName)
        {
            LblRepository.MouseOver();
            Link lnkRepository = new Link(_webDriver.FindElement(By.XPath(string.Format("//a[.='{0}']", repositoryName))));
            lnkRepository.Click();
            return new MainPage(_webDriver);
        }

        /// <summary>
        /// Verify the link is exist
        /// </summary>
        /// <param name="linkName">Name of the link.</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        public Boolean IsLinkExist(string linkName)
        {
            Link Link = new Link(_webDriver,By.XPath(string.Format("//a[.='{0}']", CommonAction.EncodeSpace(linkName))));            
            return Link.isDisplayed();
        }

        /// <summary>
        /// Determines whether [is div exist] [the specified div name].
        /// </summary>
        /// <param name="divName">Name of the div.</param>
        /// <returns></returns>
        /// <Author>Tu Nguyen</Author> 
        /// <Update>Huong Huynh-handle space</Update>
        
        public Boolean IsDivExist(string divName)
        {
            Div div = new Div(_webDriver,By.XPath(string.Format("//div[@title='{0}']", CommonAction.EncodeSpace(divName))));
            bool a = div.isExists();
            return div.isExists();
        }

        #endregion

        #region Get Text Methods
        public string GetUserNameText()
        {
            return LblUsername.Text;
        }
        public string GetCurrentRepositoryText()
        {
            return LblCurrentRepository.Text;
        }
        #endregion

        /// <summary>
        /// Switches to new opened window.
        /// </summary>
        /// <param name="driver">The driver.</param>
        /// <param name="isNewUrl">if set to <c>true</c> [is new URL].</param>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void SwitchToNewOpenedWindow(IWebDriver driver, bool isNewUrl = true)
        {
            driver.Close();
            driver.SwitchTo().Window(driver.WindowHandles.Last());
        }

        /// <summary>
        /// Waits for page load complete
        /// </summary>
        /// <return></return>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void WaitForPageLoadComplete()
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(Constants.WaitTimeoutShortSeconds));
            try
            {
                wait.Until(w => ((IJavaScriptExecutor)_webDriver).ExecuteScript("return document.readyState;").Equals("loaded"));
            }
            catch (WebDriverException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        /// <summary>
        /// Switches to new frame
        /// </summary>
        /// <param name="element">The IWebElement</param>
        /// <return></return>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void SwitchToNewFrame(IWebElement element)
        {
            _webDriver.SwitchTo().Frame(element);
        }
        /// <summary>
        /// Closes the window
        /// </summary>
        /// <return></return>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void CloseWindow()
        {
            _webDriver.Close();
        }

        /// <summary>
        /// Waits for control exists
        /// </summary>
        /// <param name="control">The By property of element</param>
        /// <param name="timeoutInSeconds">The timeout in seconds</param>
        /// <return></return>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void WaitForControlExists(By control, int timeoutInSeconds = Constants.WaitTimeoutShortSeconds)
        {
            //try
            //{
            //    _webDriver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(timeoutInSeconds));
            //    IWebElement element = _webDriver.FindElement(control);
            //    if (element == null)
            //    {
            //        WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(timeoutInSeconds));
            //        wait.Until(drv => drv.FindElement(control));
            //    }
            //}
            //catch
            //{
            //    throw new Exception("No element have been found.");
            //}
        }

        /// <summary>
        /// Refreshes the current page
        /// </summary>
        /// <return></return>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void RefreshCurrentPage()
        {
            _webDriver.Navigate().Refresh();
        }

        /// <summary>
        /// Confirm OK/YES/NO/CANCEL on the dialog
        /// </summary>
        /// <param name="buttonName">Name of the button on the dialog</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/26/2015</date>
        public void ConfirmDialog(string buttonName)
        {
            switch (buttonName.ToUpper())
            {
                case "OK":
                case "YES":
                    _webDriver.SwitchTo().Alert().Accept();
                    break;

                case "NO":
                case "CANCEL":
                    _webDriver.SwitchTo().Alert().Dismiss();
                    break;
            }
        }

        /// <summary>
        /// Get the text message in the dialog
        /// </summary>
        /// <returns>dialog message string</returns>
        /// <author>Huong Huynh</author>
        ///<date>05/26/2015</date>
        public string GetDialogText()
        {
            string dglMessage = _webDriver.SwitchTo().Alert().Text;
            return dglMessage;
        }

        /// <summary>
        /// Gets the CheckBox status.
        /// </summary>
        /// <param name="ck">The ck.</param>
        /// <returns></returns>
        /// Author: TU Nguyen
        public Boolean GetCheckBoxStatus(Checkbox ck)
        {
            return ck.Selected;
        }

        /// <summary>
        /// Gets the URL
        /// </summary>
        /// <returns>String</returns>
        /// <author>Tu Nguyen</author>
        /// <date>05/26/2015</date>
        public string GetURL()
        {
            string url = _webDriver.Url;
            return url;
        }

        public void ClickLinkText(string linkText)
        {
            Link lnkDynamic = new Link(FindElement(By.XPath(string.Format("//a[.='{0}']", CommonAction.EncodeSpace(linkText)))));
            lnkDynamic.Click();
        }

       
        public AddNewPanelPage OpenAddNewPanelPageFromButton()
        {
            BtnChoosePanel.Click();
            BtnCreateNewPanel.Click();
            return new AddNewPanelPage(_webDriver);
        }


        public bool isAlertPresent()
        {
            try
            {
                _webDriver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }

        public bool IsComboboxListed(ComboBox cbb, string[] optionList)
        {
            IList<string> cbbValues = cbb.OptionStrings;
            if (cbbValues.Count != optionList.Length)
            {
                return false;
            }
            foreach (string option in optionList)
            {
                if (!cbbValues.Contains(option))
                {
                    return false;
                }
            }
            return true;
                
        }
        #endregion
    }
}
