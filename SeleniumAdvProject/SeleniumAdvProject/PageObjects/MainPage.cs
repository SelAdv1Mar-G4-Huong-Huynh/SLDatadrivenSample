using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using SeleniumAdvProject.Common;
using SeleniumAdvProject.DataObjects;
using SeleniumAdvProject.Ultilities.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeleniumAdvProject.PageObjects
{
    public class MainPage : BasePage
    {

        #region Locators
        static readonly By _lnkDelete = By.XPath("//a[@class='delete']");
        static readonly By _divvOvelayClass = By.XPath("//div[@class='ui-dialog-overlay custom-overlay']");
        #endregion

        #region Elements
        public Link LnkDelete
        {
            get { return new Link(FindElement(_lnkDelete)); }
        }
        public Div DivOvelayClass
        {
            get { return new Div(_webDriver, _divvOvelayClass); }
            
        }
        #endregion

        #region Methods

        /// <summary>
        /// Constructor of MainPage class
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        public MainPage() { }

        /// <summary>
        /// Constructor of MainPage class
        /// </summary>
        /// <param name="webDriver">IWebDriver</param>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        public MainPage(IWebDriver webDriver) : base(webDriver) { }

        /// <summary>
        /// Delete the page
        /// </summary>
        /// <param name="pathOfPage">The path to go to the current page</param>
        /// <param name="confirmDelete">Define action Delete: click on Yes or No button, or not click on the button </param>
        /// <returns>Main Page</returns>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        public MainPage DeletePage(string pathOfPage, string confirmDelete = "Yes")
        {
            ClickDeleteLink(pathOfPage);
            ConfirmDialog(confirmDelete);
            WaitForPageLoadComplete();
            return this;
        }

        /// <summary>
        /// Clicks the delete link.
        /// </summary>
        /// <param name="pathOfPage">The path of page.</param>
        /// <returns></returns>        
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public MainPage ClickDeleteLink(string pathOfPage)
        {
            GoToPage(pathOfPage);
            LblGlobalSetting.MouseOver();
            LnkDelete.MouseOver();
            LnkDelete.Click();
            return this;
        }

        /// <summary>
        /// Adds the new panel.
        /// </summary>
        /// <param name="chart">The chart information.</param>
        /// <returns>Return main page after add new panle successfully</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public MainPage AddNewPanel(Chart chart)
        {
            GoToPage(chart.PageName);
            OpenAddNewPanelPopup().AddChart(chart).WaitForPageLoadComplete();
            return this;
        }

        /// <summary>
        /// Opens the new panel popup.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public MainPage OpenNewPanelPopUp(string page)
        {
            GoToPage(page);
            BtnChoosePanel.Click();
            // WaitForControlExists(By.XPath("//span[.='Create new panel']"), Constants.WaitTimeoutShortSeconds);
            BtnCreateNewPanel.Click();
            return this;
        }

        /// <summary>
        /// Opens the edit panel popup.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public AddNewPanelPage OpenEditPanelPopup()
        {
            LnkEditPanel.Click();
            WaitForPageLoadComplete();
            return new AddNewPanelPage(_webDriver);

        }
        /// <summary>
        /// Opens the panel configuration popup by click on Chooses Panle and click on Chart panel instance
        /// </summary>
        /// <param name="linkText">Input panle instance</param>
        /// <returns>Add New Panel Page</returns>
        /// <autho>Huong Huynh</autho>
        /// <date>06/03/2016</date>
        /// <update>Tu Nguyen</update>
        public AddNewPanelPage OpenPanelConfigurationFromChoosePanel(string linkText)
        {
            BtnChoosePanel.Click();
            WaitForPageLoadComplete();
            Link lnkDynamic = new Link(FindElement(By.XPath(string.Format("//a[.='{0}']", CommonAction.EncodeSpace(linkText)))));
            lnkDynamic.Click();
            return new AddNewPanelPage(webDriver);
        }
        /// <summary>
        /// Gets the position page.
        /// </summary>
        /// <param name="pageName">Name of the page.</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        /// <update>Tu Nguyen</update>
        public int GetPositionPage(string pageName)
        {
            string tmp = "";
            string pagename = "";
            if (pageName.Contains("Child") == true)
            {
                tmp = Regex.Replace(pageName, "ChildPage", "");
                pagename = Regex.Replace(tmp, " ", "");

            }
            else
            {
                tmp = Regex.Replace(pageName, "Page", "");
                pagename = Regex.Replace(tmp, " ", "");
            }
            Link page = new Link(FindElement(By.XPath(string.Format("//a[contains(text(),'{0}')]", pagename))));
            return page.Location.X;

        }

        /// <summary>
        /// Gets the column count.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public int GetColumnCount()
        {
            IList<IWebElement> elements = _webDriver.FindElements(By.XPath("//ul[@class='column ui-sortable']"));
            int columnCount = elements.Count;
            return columnCount;
        }

        /// <summary>
        /// Determines whether [is page existed] [the specified page name].
        /// </summary>
        /// <param name="pageName">Name of the page.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public Boolean IsPageExisted(string pageName)
        {
            string tmp = "";
            string pagename = "";
            if (pageName.Contains("Child") == true)
            {
                tmp = Regex.Replace(pageName, "ChildPage", "");
                pagename = Regex.Replace(tmp, " ", "");

            }
            else
            {
                tmp = Regex.Replace(pageName, "Page", "");
                pagename = Regex.Replace(tmp, " ", "");
            }
            Link lnkpage = new Link(FindElement(By.XPath(string.Format("//a[contains(text(),'{0}')]", pagename))));
            return lnkpage.Enabled;
        }

        /// <summary>
        /// Opens the setting.
        /// </summary>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public MainPage OpenSetting()
        {
            LblGlobalSetting.MouseOver();
            return this;
        }

        /// <summary>
        /// Verify Global Setting icon is exist
        /// </summary>
        /// <returns>True/False</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public Boolean IsSettingExist()
        {
            return LblGlobalSetting.isDisplayed();
        }

        /// <summary>
        /// Verify Page is exist
        /// </summary>
        /// <param name="pathOfPage">The path to go to the current page</param>
        /// <returns>True/False</returns>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        public Boolean IsPageExist(string pathOfPage)
        {
            string[] arrNode = pathOfPage.Split('/');
            string finalNode = arrNode[arrNode.Length - 1];
            if (arrNode.Length == 1)
                return IsLinkExist(arrNode[0]);

            GoToLink(pathOfPage.Substring(0, pathOfPage.Length - finalNode.Length - 1));
            string xpathOfPage = string.Format("//a[.='{0}']", CommonAction.EncodeSpace(arrNode[0]));
            for (int i = 1; i < arrNode.Length; i++)
            {
                xpathOfPage += string.Format("/..//a[.='{0}']", CommonAction.EncodeSpace(arrNode[i]));
            }
            Link LnkPage = new Link(_webDriver, By.XPath(xpathOfPage));
            bool a = LnkPage.isDisplayed();
            return LnkPage.isDisplayed();
        }

        /// <summary>
        /// Determines whether [is page display after] [the specified page name1].
        /// </summary>
        /// <param name="pageName1">The page name1.</param>
        /// <param name="pageName2">The page name2.</param>
        /// <returns>return true if page exist and false if not exist</returns>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public bool IsPageDisplayAfter(string pageName1, string pageName2)
        {
            Label pageTab = new Label(FindElement(By.XPath(string.Format("//div[@id='main-menu']/div/ul/li[.='{0}']/preceding-sibling::li[1]",CommonAction.EncodeSpace(pageName2)
                ,CommonAction.EncodeSpace(pageName1)))));
            string tempPage = pageTab.Text;
            return tempPage.Equals(pageName1);
        }

        /// <summary>
        /// Verify Main Page is displayed
        /// </summary>
        /// <returns>True/False</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2015</date>
        /// <Update>Tu Nguyen</Update>
        public bool Displayed(string username)
        {
            Link lnkWelcome = new Link(FindElement(By.XPath(string.Format("//a[.='{0}' and @href='#Welcome']", username))));
            return lnkWelcome.Enabled;
        }

        /// <summary>
        /// Add the new page
        /// </summary>
        /// <param name="page">The page object</param>
        /// <returns>The MainPage object</returns>
        /// <author>Vu Tran</author>
        /// <date>05/26/2015</date>
        public MainPage AddPage(Page page)
        {
            OpenAddNewPage().AddPage(page.ParentPage, page);
            return this;
        }

        /// <summary>
        /// Adds the page with error.
        /// </summary>
        /// <param name="page">The page.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string AddPageWithError(Page page)
        {
            return OpenAddNewPage().AddPageWithExpectedError(page.ParentPage, page);

        }

        /// <summary>
        /// Edits the page.
        /// </summary>
        /// <param name="page">The page.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public MainPage EditPage(Page page)
        {
            AddNewPage newPage = new AddNewPage(_webDriver);
            newPage.EditPage(page.ParentPage, page);
            return this;
        }
        
        /// <summary>
        /// Determines the content in a table is sorted or not?
        /// </summary>
        /// <param name="combobox">the name of table</param>
        /// <param name="sortType">Type of the sort DESC|ASC.</param>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        public bool IsContentInTableSorted(string tableName, string sortType)
        {
            Table table = new Table(FindElement(By.XPath(string.Format("//div[.='{0}']/../table", tableName))));
            IList<IWebElement> rows = table.FindElements(By.XPath(string.Format("//div[.='{0}']/../table/tbody/tr", tableName)));
            List<string> tableContent = new List<string>();
            for (int i = 0; i < rows.Count(); i++)
            {
                IList<IWebElement> columns = table.FindElements(By.XPath(string.Format("//div[.='{0}']/../table/tbody/tr[{1}]/td", tableName, i + 1)));
                foreach (IWebElement column in columns)
                {
                    tableContent.Add(column.Text);
                }

            }

            bool flag = false;
            if (tableContent.Count == 1)
            {
                flag = true;
            }
            else
            {
                for (int i = 1; i < tableContent.Count - 1; i++)
                {
                    if (sortType == "DESC")
                    {
                        if (tableContent[i].CompareTo(tableContent[i + 1]) >= 0)
                            flag = true;
                    }
                    else if (sortType == "ASC")
                    {
                        if (tableContent[i].CompareTo(tableContent[i + 1]) <= 0)
                            flag = true;
                    }
                }
            }
            return flag;
        }

        public AddNewPanelPage ClickEditPanelIcon(string panelName)
        {
            Button iconEditPanel = new Button(FindElement(By.XPath(string.Format("//div[@title='{0}']/../following-sibling::div//li[@title='Edit Panel']", CommonAction.EncodeSpace(panelName)))));
            iconEditPanel.Click();
            return new AddNewPanelPage(_webDriver);
        }

        #endregion
    }
}
