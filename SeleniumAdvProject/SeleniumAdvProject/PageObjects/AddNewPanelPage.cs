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
    public class AddNewPanelPage : Popup
    {
        #region Locators

        static readonly By _cbbDataProfile = By.XPath("//select [@id='cbbProfile']");
        static readonly By _txtDisplayName = By.XPath("//input[@id='txtDisplayName']");
        static readonly By _txtChartTitle = By.XPath("//input[@id='txtChartTitle']");
        static readonly By _chbShowTitle = By.XPath("//input[@id='chkShowTitle']");
        static readonly By _cbbChartType = By.XPath("//select[@id='cbbChartType']");
        static readonly By _cbbCategory = By.XPath("//select[@id='cbbCategoryField']");
        static readonly By _cbbSeries = By.XPath("//select[@id='cbbSeriesField']");
        static readonly By _txtCategoryCaption = By.XPath("//input[@id='txtCategoryXAxis']");
        static readonly By _txtSeriesCaption = By.XPath("//input[@id='txtValueYAxis']");
        static readonly By _chbSeries = By.XPath("//input[@id='chkSeriesName']");
        static readonly By _chbCategories = By.XPath("//input[@id='chkCategoriesName']");
        static readonly By _chbValue = By.XPath("//input[@id='chkValue']");
        static readonly By _chbPercentage = By.XPath("//input[@id='chkPercentage']");
        static readonly By _rbChart = By.XPath("//input[@id='radPanelType0']");
        static readonly By _rbIndicator = By.XPath("//input[@id='radPanelType1']");
        static readonly By _rbReport = By.XPath("//input[@id='radPanelType2']");
        static readonly By _rbHeadMap = By.XPath("//input[@id='radPanelType3']");
        static readonly By _rb2D = By.XPath("//input[@id='rdoChartStyle2D']");
        static readonly By _rb3D = By.XPath("//input[@id='rdoChartStyle3D']");
        static readonly By _rbNone = By.XPath("//input[@id='radPlacementNone']");
        static readonly By _rbTop = By.XPath("//input[@id='radPlacementTop']");
        static readonly By _rbRight = By.XPath("//input[@id='radPlacementRight']");
        static readonly By _rbBottom = By.XPath("//input[@id='radPlacementBottom']");
        static readonly By _rlLeft = By.XPath("//input[@id='radPlacementLeft']");
        static readonly By _cbbSelectPage = By.XPath("//select[@id='cbbPages']");
        static readonly By _txtHeight = By.XPath("//input[@id='txtHeight']");
        static readonly By _txtFolder = By.XPath("//input[@id='txtFolder']");
        static readonly By _btnOKConfigurationPanel = By.XPath(" //input[@id='OK' and contains(@onclick,'Dashboard.addPanelToPage')]");
        static readonly By _btnCancelConfigurationPanel = By.XPath(" //input[@id='OK' and contains(@onclick,'Dashboard.closePanelDialog')]");
        static readonly By _btnFolder = By.XPath("//img[@class='panel_setting_treefolder']");
        static readonly By _btnFolderSelectionOK = By.XPath("//input[@id='btnFolderSelectionOK']");

        #endregion

        #region Elements
        public ComboBox CbbDataProfile
        {
            get { return new ComboBox(FindElement(_cbbDataProfile)); }
        }
        public Button BtnCancelConfigurationPanel
        {
            get { return new Button(FindElement(_btnCancelConfigurationPanel)); }
        }
        public Button BtnFolder
        {
            get { return new Button(FindElement(_btnFolder)); }
        }
        public Button BtnFolderSelectionOK
        {
            get { return new Button(FindElement(_btnFolderSelectionOK)); }
        }
        public TextBox TxtDisplayName
        {
            get { return new TextBox(FindElement(_txtDisplayName)); }
        }
        public TextBox TxtChartTitle
        {
            get { return new TextBox(FindElement(_txtChartTitle)); }
        }
        public Checkbox ChbShowTitle
        {
            get { return new Checkbox(FindElement(_chbShowTitle)); }
        }
        public ComboBox CbbChartType
        {
            get { return new ComboBox(FindElement(_cbbChartType)); }
        }
        public RadioButton RbChart
        {
            get { return new RadioButton(FindElement(_rbChart)); }
        }
        public RadioButton RbReport
        {
            get { return new RadioButton(FindElement(_rbReport)); }
        }
        public RadioButton RbIndicator
        {
            get { return new RadioButton(FindElement(_rbIndicator)); }
        }
        public RadioButton RbHeadMap
        {
            get { return new RadioButton(FindElement(_rbHeadMap)); }
        }

        public ComboBox CbbCategory
        {
            get { return new ComboBox(FindElement(_cbbCategory)); }
        }
        public ComboBox CbbSeries
        {
            get { return new ComboBox(FindElement(_cbbSeries)); }
        }
        public RadioButton Rb2D
        {
            get { return new RadioButton(FindElement(_rb2D)); }
        }
        public RadioButton Rb3D
        {
            get { return new RadioButton(FindElement(_rb3D)); }
        }

        public TextBox TxtCategoryCaption
        {
            get { return new TextBox(FindElement(_txtCategoryCaption)); }
        }
        public TextBox TxtSeriesCaption
        {
            get { return new TextBox(FindElement(_txtSeriesCaption)); }
        }
        public Checkbox ChbSeries
        {
            get { return new Checkbox(FindElement(_chbSeries)); }
        }
        public Checkbox ChbCategories
        {
            get { return new Checkbox(FindElement(_chbCategories)); }
        }
        public Checkbox ChbValue
        {
            get { return new Checkbox(FindElement(_chbValue)); }
        }
        public Checkbox ChbPercentage
        {
            get { return new Checkbox(FindElement(_chbPercentage)); }
        }
        public RadioButton RbNone
        {
            get { return new RadioButton(FindElement(_rbNone)); }
        }
        public RadioButton RbTop
        {
            get { return new RadioButton(FindElement(_rbTop)); }
        }
        public RadioButton RbRight
        {
            get { return new RadioButton(FindElement(_rbRight)); }
        }
        public RadioButton RbBottom
        {
            get { return new RadioButton(FindElement(_rbBottom)); }
        }
        public RadioButton RbLeft
        {
            get { return new RadioButton(FindElement(_rlLeft)); }
        }
        public ComboBox CbbSelectPage
        {
            get { return new ComboBox(FindElement(_cbbSelectPage)); }
        }
        public TextBox TxtHeight
        {
            get { return new TextBox(FindElement(_txtHeight)); }
        }
        public RadioButton TxtFolder
        {
            get { return new RadioButton(FindElement(_txtFolder)); }
        }
        public Button BtnOKConfigurationPanel
        {
            get { return new Button(FindElement(_btnOKConfigurationPanel)); }
        }
        #endregion

        #region Methods
        public AddNewPanelPage() { }
        public AddNewPanelPage(IWebDriver webDriver) : base(webDriver) { }

        #region Private Methods

        #region Private Methods

        public bool IsDataProfileExists(string dataProfileName)
        {
            IList<string> dataProfiles = CbbDataProfile.OptionStrings;
            return dataProfiles.Contains(dataProfileName);
        }

        /// <summary>
        /// Selects the legend.
        /// </summary>
        /// <param name="legend">The legend.</param>
        public void SelectLegend(string legend)
        {
            switch (legend)
            {
                case "Top":
                    RbTop.Click();
                    break;
                case "Right":
                    RbRight.Click();
                    break;
                case "Button":
                    RbRight.Click();
                    break;
                case "Left":
                    RbLeft.Click();
                    break;
                default:
                    RbNone.Click();
                    break;
            }
        }

        /// <summary>
        /// Gets the legend.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetLegend()
        {
            string legend = "";
            string tmp1 = RbNone.GetAttribute("checked");
            if (tmp1 == "true")
            {
                legend = "None";
            }
            string tmp2 = RbLeft.GetAttribute("checked");
            if (tmp2 == "true")
            {
                legend = "Left";
            }
            string tmp3 = RbRight.GetAttribute("checked");
            if (tmp3 == "true")
            {
                legend = "Right";
            }
            string tmp4 = RbTop.GetAttribute("checked");
            if (tmp4 == "true")
            {
                legend = "Top";
            }
            string tmp5 = RbBottom.GetAttribute("checked");
            if (tmp5 == "true")
            {
                legend = "Bottom";
            }
            return legend;
        }

        /// <summary>
        /// Gets the style.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetStyle()
        {
            string style = "";
            string tmp1 = Rb2D.GetAttribute("checked");
            if (tmp1 == "true")
            {
                style = "2D";
            }
            string tmp2 = Rb3D.GetAttribute("checked");
            if (tmp2 == "true")
            {
                style = "3D";
            }
            return style;

        }

        /// <summary>
        /// Gets the type.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetPanelType()
        {
            string type = "";
            string tmp1 = RbChart.GetAttribute("checked");
            if (tmp1 == "true")
            {
                type = "Chart";
            }
            string tmp2 = RbIndicator.GetAttribute("checked");
            if (tmp2 == "true")
            {
                type = "Indicator";
            }
            string tmp3 = RbReport.GetAttribute("checked");
            if (tmp3 == "true")
            {
                type = "Report";
            }
            string tmp4 = RbHeadMap.GetAttribute("checked");
            if (tmp4 == "true")
            {
                type = "Heat Map";
            }
            return type;
        }

        /// <summary>
        /// Selects the type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// Author: Tu Nguyen
        private void SelectType(string type)
        {
            switch (type)
            {
                case "Chart":
                    RbChart.Click();
                    break;
                case "Indicator":
                    RbIndicator.Click();
                    break;
                case "Report":
                    RbReport.Click();
                    break;
                case "Heat Map":
                    RbHeadMap.Click();
                    break;
            }
        }

        /// <summary>
        /// Fills all panel data.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="dataProfile">The data profile.</param>
        /// <param name="displayName">The display name.</param>
        /// <param name="chartTitle">The chart title.</param>
        /// <param name="showTitle">The show title.</param>
        /// <param name="chartType">Type of the chart.</param>
        /// <param name="style">The style.</param>
        /// <param name="category">The category.</param>
        /// <param name="series">The series.</param>
        /// <param name="dataLabel">The data label.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public AddNewPanelPage FillPanelData(string type, string dataProfile, string displayName, string chartTitle, string showTitle, string chartType, string style, string category, string series, string dataLabel, bool selected, string legend)
        {
            SelectType(type);
            SelectChartType(chartType);
            CbbCategory.SelectByTextFromGroup(category);
            CbbSeries.SelectByTextFromGroup(series);
            TxtDisplayName.EnterText(displayName);
            TxtChartTitle.EnterText(chartTitle);
            CbbDataProfile.SelectByTextFromGroup(dataProfile);
            bool temp = GetCheckBoxStatus(ChbShowTitle);
            switch (showTitle)
            {
                case "on":
                    if (temp == false) { ChbShowTitle.Click(); }
                    break;
                case "off":
                    if (temp == true) { ChbShowTitle.Click(); }
                    break;
            }
            SelectStyle(style);
            SelectDataLabel(dataLabel, selected);
            SelectLegend(legend);
            return this;
        }



        /// <summary>
        /// Selects the data labels.
        /// </summary>
        /// <param name="dataLabel">The data label.</param>
        /// 
        private void SelectDataLabels(string[] dataLabel)
        {
            if ((dataLabel == null))
                return;
            for (int i = 0; i < dataLabel.Length; i++)
            {
                if (dataLabel[i].Equals("Series"))
                    ChbSeries.Click();
                if (dataLabel[i].Equals("Categories"))
                    ChbCategories.Click();
                if (dataLabel[i].Equals("Value"))
                    ChbValue.Click();
                if (dataLabel[i].Equals("Percentage"))
                    ChbPercentage.Click();
            }
        }

        /// <summary>
        /// Selects the data label.
        /// </summary>
        /// <param name="dataLabel">The data label.</param>
        /// Author: TuNguyen
        private void SelectDataLabel(string dataLabel, bool selected)
        {
            if ((dataLabel == null))
                return;
            switch (dataLabel)
            {
                case "Series":
                    if (selected == true)
                    {
                        ChbSeries.Check();
                    }
                    else
                    {
                        ChbSeries.Uncheck();
                    }
                    break;
                case "Categories":
                    if (selected == true)
                    {
                        ChbCategories.Check();
                    }
                    else
                    {
                        ChbCategories.Uncheck();
                    }
                    break;
                case "Value":
                    if (selected == true)
                    {
                        ChbValue.Check();
                    }
                    else
                    {
                        ChbValue.Uncheck();
                    }
                    break;
                case "Percentage":
                    if (selected == true)
                    {
                        ChbPercentage.Check();
                    }
                    else
                    {
                        ChbPercentage.Uncheck();
                    }
                    break;
            }
        }
        /// <summary>
        /// Selects the style.
        /// </summary>
        /// <param name="style">The style.</param>
        public void SelectStyle(string style)
        {
            switch (style)
            {
                case "3D":
                    Rb3D.Click();
                    break;
                default:
                    break;
            }
        }


        /// <summary>
        /// Gets the category status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetCategoryStatus()
        {
            string status = "enabled";
            string temp = CbbCategory.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the category caption status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetCategoryCaptionStatus()
        {
            string status = "enabled";
            string temp = TxtCategoryCaption.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the Series status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetSeriesStatus()
        {
            string status = "enabled";
            string temp = CbbSeries.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the series caption status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetSeriesCaptionStatus()
        {
            string status = "enabled";
            string temp = TxtSeriesCaption.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the categories CheckBox status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetCategoriesCheckBoxStatus()
        {
            string status = "enabled";
            string temp = ChbCategories.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the series CheckBox status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetSeriesCheckBoxStatus()
        {
            string status = "enabled";
            string temp = ChbSeries.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the value CheckBox status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetValueCheckBoxStatus()
        {
            string status = "enabled";
            string temp = ChbValue.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Gets the percentage CheckBox status.
        /// </summary>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public string GetPercentageCheckBoxStatus()
        {
            string status = "enabled";
            string temp = ChbPercentage.GetAttribute("disabled");
            switch (temp)
            {
                case "true":
                    status = "disabled";
                    break;
                case "false":
                    status = "enabled";
                    break;
            }
            return status;
        }
        /// <summary>
        /// Selects the type of the chart.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        /// Author: Tu Nguyen
        public void SelectChartType(string type)
        {

            switch (type)
            {
                case "Pie":
                    CbbChartType.SelectByIndex(0);
                    break;
                case "Single Bar":
                    CbbChartType.SelectByIndex(1);
                    break;
                case "Stacked Bar":
                    CbbChartType.SelectByIndex(2);
                    break;
                case "Group Bar":
                    CbbChartType.SelectByIndex(3);
                    break;
                case "Line":
                    CbbChartType.SelectByIndex(4);
                    break;
            }

        }

        /// <summary>
        /// Closes the panel dialog.
        /// </summary>
        /// Author: Tu Nguyen
        public void ClosePanelDialog(string button)
        {
            switch (button)
            {
                case "OK":
                    BtnOk.Click();
                    break;
                case "Cancel":
                    BtnCancel.Click();
                    break;
            }
        }

        #endregion
        /// <summary>
        /// Adds the chart.
        /// </summary>
        /// <param name="pChart">Chart object.</param>
        /// <returns>Main Page</returns>
        /// <author>Huong Huynh</author>
        /// <date>5/30/2016</date>
        /// <param name="pChart">The p chart.</param>
        /// <returns></returns>
        /// Update: Tu Nguyen
        public MainPage AddChart(Chart pChart)
        {
            RbChart.Click();
            TxtDisplayName.SendKeys(pChart.DisplayName);
            TxtChartTitle.SendKeys(pChart.ChartTitle);
            CbbDataProfile.SelectByTextFromGroup(pChart.DataProfile);
            bool temp = GetCheckBoxStatus(ChbShowTitle);
            switch (pChart.ShowTitle)
            {
                case true:
                    if (temp == false) { ChbShowTitle.Click(); }
                    break;
                case false:
                    if (temp == true) { ChbShowTitle.Click(); }
                    break;
            }
            SelectChartType(pChart.ChartType);
            CbbCategory.SelectByTextFromGroup(pChart.Category);
            TxtCategoryCaption.SendKeys(pChart.CategoryCaption);
            CbbSeries.SelectByTextFromGroup(pChart.Series);
            TxtSeriesCaption.SendKeys(pChart.SeriesCaption);
            SelectLegend(pChart.Legend);
            SelectDataLabels(pChart.DataLabel);
            SelectStyle(pChart.Style);
            BtnOk.Click();
            SettingPanel(pChart.PageName, pChart.Height, pChart.Folder);
            return new MainPage(_webDriver);
        }
        public void EditChartPanel(Chart pChart)
        {
            TxtDisplayName.SendKeys(pChart.DisplayName);
            TxtChartTitle.SendKeys(pChart.ChartTitle);
            CbbDataProfile.SelectByTextFromGroup(pChart.DataProfile);
            bool temp = GetCheckBoxStatus(ChbShowTitle);
            switch (pChart.ShowTitle)
            {
                case true:
                    if (temp == false) { ChbShowTitle.Click(); }
                    break;
                case false:
                    if (temp == true) { ChbShowTitle.Click(); }
                    break;
            }
            SelectChartType(pChart.ChartType);
            CbbCategory.SelectByTextFromGroup(pChart.Category);
            TxtCategoryCaption.SendKeys(pChart.CategoryCaption);
            CbbSeries.SelectByTextFromGroup(pChart.Series);
            TxtSeriesCaption.SendKeys(pChart.SeriesCaption);
            SelectLegend(pChart.Legend);
            SelectDataLabels(pChart.DataLabel);
            SelectStyle(pChart.Style);
            BtnOk.Click();
            SettingPanel(pChart.PageName, pChart.Height, pChart.Folder);
        }

        public string AddChartWithExpectedError(Chart pChart)
        {
            RbChart.Click();
            TxtDisplayName.SendKeys(pChart.DisplayName);
            TxtChartTitle.SendKeys(pChart.ChartTitle);
            CbbDataProfile.SelectByTextFromGroup(pChart.DataProfile);
            bool temp = GetCheckBoxStatus(ChbShowTitle);
            switch (pChart.ShowTitle)
            {
                case true:
                    if (temp == false) { ChbShowTitle.Click(); }
                    break;
                case false:
                    if (temp == true) { ChbShowTitle.Click(); }
                    break;
            }
            SelectChartType(pChart.ChartType);
            CbbCategory.SelectByTextFromGroup(pChart.Category);
            TxtCategoryCaption.SendKeys(pChart.CategoryCaption);
            CbbSeries.SelectByTextFromGroup(pChart.Series);
            TxtSeriesCaption.SendKeys(pChart.SeriesCaption);
            SelectLegend(pChart.Legend);
            SelectDataLabels(pChart.DataLabel);
            SelectStyle(pChart.Style);
            BtnOk.Click();
            SettingPanel(pChart.PageName, pChart.Height, pChart.Folder);
            string errorMesage = GetDialogText();
            ConfirmDialog("OK");
            return errorMesage;
        }

        /// <summary>
        /// Settings the panel.
        /// </summary>
        /// <param name="selectPage">The select page.</param>
        /// <param name="height">The height.</param>
        /// <param name="folder">The folder.</param>
        /// <author>Huong Huynh</author>
        /// <date>5/30/2016</date>
        public MainPage SettingPanel(string selectPage, int height, string folder)
        {
            if (selectPage != null)
            {
                CbbSelectPage.SelectByText(selectPage);
                TxtHeight.SendKeys(height.ToString());
                TxtFolder.SendKeys(folder);
                BtnOKConfigurationPanel.Click();
                WaitForPageLoadComplete();
            }
            return new MainPage(_webDriver);
        }

        public void SelectFolder(string folderPath)
        {
            BtnFolder.Click();
            string[] arrNode = folderPath.Split('/');
            for (int i = 1; i < arrNode.Length; i++)
            {
                IWebElement plusImg = FindElement(By.XPath(string.Format("//div[@id='async_html_2']//a[text()=' {0}']//preceding-sibling::a/img[@src='images/plus.gif']", arrNode[i])));
                if (plusImg != null)
                {
                    plusImg.Click();
                }
            }
            IWebElement LnkFinalNode = FindElement(By.XPath(string.Format("//div[@id='async_html_2']//a[text()=' {0}']", arrNode[arrNode.Length - 1])));
            LnkFinalNode.Click();
            BtnFolderSelectionOK.Click();

        }

        public bool IsSelectFolderPopulation(string folderPath)
        {
            BtnFolder.Click();
            string[] arrNode = folderPath.Split('/');
            for (int i = 0; i < arrNode.Length; i++)
            {
                IWebElement plusImg = FindElement(By.XPath(string.Format("//div[@id='async_html_2']//a[text()=' {0}']//preceding-sibling::a/img[@src='images/plus.gif']", arrNode[i])));
                if (plusImg != null)
                {
                    return false;
                }
            }
            return true;

        }

        public AddNewPanelPage FillPanelConfiguration(string selectPage, string height, string folder)
        {
            CbbSelectPage.SelectByText(selectPage);
            TxtHeight.SendKeys(height);
            TxtFolder.SendKeys(folder);
            return this;
        }



        //public string SettingPanelWithExpectedError(string selectPage, int height, string folder)
        //{
        //    CbbSelectPage.SelectByText(selectPage);
        //    TxtHeight.SendKeys(height.ToString());
        //    TxtFolder.SendKeys(folder);
        //    BtnOKConfigurationPanel.Click();
        //    return this.GetDialogText();
        //}
        //public string SettingPanelWithExpectedError(string selectPage, double height, string folder)
        //{
        //    CbbSelectPage.SelectByText(selectPage);
        //    TxtHeight.SendKeys(height.ToString());
        //    TxtFolder.SendKeys(folder);
        //    BtnOKConfigurationPanel.Click();
        //    return this.GetDialogText();
        //}
        /// <summary>
        /// Settings the panel with expected error.
        /// </summary>
        /// <param name="selectPage">select page name, left defaut if this value is null.</param>
        /// <param name="height">enter height, left as default if this value is null</param>
        /// <param name="folder">enter folder, left as default if this value is null.</param>
        /// <returns></returns>
        /// <author>Huong Huynh</author>
        /// <date>6/3/2016</date>
        public string SettingPanelWithExpectedError(string selectPage, object height, string folder)
        {
            CbbSelectPage.SelectByText(selectPage);
            if (height != null)
            {
                TxtHeight.SendKeys(height.ToString());
            }
            TxtFolder.SendKeys(folder);
            BtnOKConfigurationPanel.Click();
            return this.GetDialogText();
        }

        /// <summary>
        /// Determines whether [is the list is sorted] [the specified combobox].
        /// </summary>
        /// <param name="combobox">The combobox.</param>
        /// <param name="sortType">Type of the sort.</param>
        /// <returns>return true the value in combobox are sorted, otherwise return false </returns>
        /// <author>Huong Huynh</author>
        /// <date>6/3/2016</date>
        public bool IsTheListIsSorted(ComboBox combobox, string sortType)
        {
            IList<string> listValues = combobox.OptionStrings;

            int rowCount = listValues.Count;
            bool flag = false;

            // start from 1 to skip the table header row run to 'i < rowCount - 1' because we check
            // a pair of row at a time
            for (int i = 1; i < rowCount - 1; i++)
            {
                if (sortType == "DESC")
                {
                    if (listValues[i].CompareTo(listValues[i + 1]) >= 0)
                        flag = true;
                }
                else if (sortType == "ASC")
                {
                    if (listValues[i].CompareTo(listValues[i + 1]) <= 0)
                        flag = true;
                }
            }
            return flag;
        }


        #endregion
        #endregion


    }
}
