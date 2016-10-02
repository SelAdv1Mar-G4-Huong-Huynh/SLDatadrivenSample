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
using System.Collections.ObjectModel;

namespace SeleniumAdvProject.PageObjects
{
    public class DataProfilePage : DataProfileBasePage
    {
        #region Locators
        static readonly By _lnkAddNew = By.XPath("//a[.='Add New']");
        static readonly By _lnkDelete = By.XPath("//a[.='Delete']");
        static readonly By _lnkCheckAll = By.XPath("//a[.='Check All']");
        static readonly By _lnkUnCheckAll = By.XPath("//a[.='UnCheck All']");
        #endregion

        #region Elements
        public Link LnkAddNew
        {
            get { return new Link(FindElement(_lnkAddNew)); }
        }
        public Link LnkDelete
        {
            get { return new Link(FindElement(_lnkDelete)); }
        }
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
        public DataProfilePage() : base() { }

        public DataProfilePage(IWebDriver webDriver) : base(webDriver) { }

        /// <summary>
        /// Delete all Data Profiles
        /// </summary>
        /// <returns>DataProfilePage object</returns>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        public DataProfilePage DeleteAllDataProfiles()
        {
            LnkCheckAll.Click();
            LnkDelete.Click();
            ConfirmDialog("OK");
            return this;
        }

        /// <summary>
        /// Deletes the data profiles.
        /// </summary>
        /// <param name="profileName">Name of the profile.</param>
        /// <returns> Data Profile Page</returns>
        /// <author>Huong Huynh</author>
        /// <date>06/10/2016</date>
        public DataProfilePage DeleteDataProfiles(string profileName)
        {
            Checkbox chkprofile = new Checkbox(FindElement(By.XPath(string.Format("//td[.='{0}']//preceding-sibling::td/input[@name='chkDel']", profileName))));
            chkprofile.Check();
            Link lnkEdit = new Link(FindElement(By.XPath(string.Format("//td[.='{0}']//following-sibling::td/a[.='Delete']", profileName))));
            lnkEdit.Click();
            ConfirmDialog("OK");
            return this;
        }
        /// <summary>
        /// Go to General Setting Page
        /// </summary>
        /// <returns>General Setting Page</returns>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        public GeneralSettingsPage GoToGeneralSettingPage()
        {
            LnkAddNew.Click();
            return new GeneralSettingsPage(_webDriver);
        }

        /// <summary>
        /// Verify Pre-set Data Profile are populated
        /// </summary>
        /// <returns>True/False</returns>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        public bool IsPresetDataProfilePopulated()
        {
            for (int i = 0; i < Constants.presetDataProfile.Length; i++)
            {
                IWebElement element = FindElement(By.XPath(string.Format("//form[@id='form1']//table/tbody//td[text()='{0}']", CommonAction.EncodeSpace(Constants.presetDataProfile[i]))));
                if (element == null)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Verify the link exist in the Data Profile table
        /// </summary>
        /// <returns>True/False</returns>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        public bool IsLinkExists(string dataProfileName, string linkName)
        {
            Link lnkName = new Link(FindElement(By.XPath(string.Format("//td[.='{0}']//following-sibling::td/a[.='{1}']", CommonAction.EncodeSpace(dataProfileName), linkName))));

            if (lnkName == null)
                return false;
            return true;
        }

        /// <summary>
        /// Determines whether [is CheckBox exists] [the specified data profile name].
        /// </summary>
        /// <param name="dataProfileName">Name of the data profile.</param>
        /// <returns></returns>
        /// Author: Vu Tran
        /// Update: Tu Nguyen
        public bool IsCheckBoxExists(string dataProfileName)
        {
            Checkbox chkDataProfile = new Checkbox(FindElement(By.XPath(string.Format("//td[.='{0}']//preceding-sibling::td[@class='chkCol']", CommonAction.EncodeSpace(dataProfileName)))));
            if (chkDataProfile == null)
                return false;
            return true;
        }

        /// <summary>
        /// Determines whether [is data profile content sorted] [the specified table name].
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="sortType">Type of the sort.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public bool IsDataProfileContentSorted(string tableName, string sortType)
        {
            Table table = new Table(FindElement(By.XPath(string.Format("//div[@class='{0}']//table", tableName))));
            IList<IWebElement> rows = table.FindElements(By.XPath(string.Format("//div[@class='{0}']//table/tbody/tr", tableName)));
            List<string> tableContent = new List<string>();
            for (int i = 0; i < rows.Count(); i++)
            {
                foreach (IWebElement row in rows)
                {
                    tableContent.Add(row.Text);
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
                        if (tableContent[i].CompareTo(tableContent[i + 1]) > 0)
                            flag = true;
                    }
                    else if (sortType == "ASC")
                    {
                        if (tableContent[i].CompareTo(tableContent[i + 1]) < 0)
                            flag = true;
                    }
                }
            }
            return flag;
        }
        
        #endregion
    }
}
