using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumAdvProject.PageObjects;
using SeleniumAdvProject.Common;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using System.Text.RegularExpressions;
using SeleniumAdvProject.DataObjects;
using System.Collections.Generic;

namespace SeleniumAdvProject.TestCases
{
    [TestClass]
    public class PanelTestCase : BaseTestCase
    {
        /// <summary>
        /// Verify that when "Choose panels" form is expanded all pre-set panels are populated and sorted correctly
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC027()
        {
            Console.WriteLine("DA_PANEL_TC027 - Verify that when \"Choose panels\" form is expanded all pre-set panels are populated and sorted correctly");

            //Set variables
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", page1.PageName);

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Go to Global Setting -> Add page
            //4 Enter page name to Page Name field.
            //5 Click OK button
            //6 Go to Global Setting -> Create Panel     
            //7 Enter display name into Display Name textbox
            //8 Select any value in Series dropdown list
            //9 Click OK button on PanelConfiguration popup and then click on Chose Panel menu icon
            LoginPage loginPage = new LoginPage(_webDriver)
                .Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(page1)
                .AddNewPanel(chart);
            mainPage.BtnChoosePanel.Click();

            ////VP Verify that all pre-set panels are populated and sorted correctly
            Assert.AreEqual(true, mainPage.IsContentInTableSorted("Charts", "ASC"), "Contents in Chart table are not sorted by ASC correctly");
            Assert.AreEqual(true, mainPage.IsContentInTableSorted("Reports", "ASC"), "Contents in Reports table are not sorted by ASC correctly");
            Assert.AreEqual(true, mainPage.IsContentInTableSorted("Indicators", "ASC"), "Contents in Indicators table are not sorted by ASC correctly");

            //Post-Condition
            //Logout			
            //Close Dashboard      
            mainPage.DeletePage(page1.PageName, "Yes")
                .OpenPanelsPage()
                .DeletePanels(chart.DisplayName)
                .Logout();
        }

        /// <summary>
        /// Verify that when "Add New Panel" form is on focused all other control/form is disabled or locked
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC028()
        {
            Console.WriteLine("DA_PANEL_TC028 - Verify that when \"Add New Panel\" form is on focused all other control/form is disabled or locked.");

            //1 Navigate to Dashboard login page
            //2 Login with valid account

            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Try to click other controls when Add New Panel dialog is opening            

            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.OpenAddNewPanelPopupFromLink();

            //VP All control/form are disabled or locked when Add New Panel dialog is opening            
            Assert.AreEqual(mainPage.DivOvelayClass.isExists(), true, "Controls are disabled");
        }

        /// <summary>
        /// Verify that user is unable to create new panel when (*) required field is not filled
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC029()
        {
            Console.WriteLine("DA_PANEL_TC029 - Verify that user is unable to create new panel when (*) required field is not filled");

            //1 Navigate to Dashboard login page
            //2 Select specific repository
            //3 Enter valid username and password
            //4 Click on Login button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3 Click Administer/Panels link            
            //5 Click Add New link
            //6 Click on OK button    
            Chart chart = new Chart();
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.OpenAddNewPanelPopupFromLink().AddChart(chart);

            //VP Warning message: "Display Name is required field" show up
            Assert.AreEqual("Display Name is a required field.", mainPage.GetDialogText(), "Failed! Actual message: {0}", mainPage.GetDialogText());
        }

        /// <summary>
        /// Verify that no special character except '@' character is allowed to be inputted into "Display Name" field
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC030()
        {
            Console.WriteLine("DA_PANEL_TC030 - Verify that no special character except '@' character is allowed to be inputted into \"Display Name\" field");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Enter value into Display Name field with special characters except "@"
            //7 Click on OK button 
            Chart chart = new Chart("Logigear#$%", "Name", null);
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            AddNewPanelPage addPanelPopup = panelPage.OpenAddNewPanelPopupFromLink();
            addPanelPopup.AddChart(chart);

            //VP Message "Invalid display name. The name can't contain high ASCII characters or any of following characters: /:*?<>|"#{[]{};" is displayed
            Assert.AreEqual("Invalid display name. The name cannot contain high ASCII characters or any of the following characters: /:*?<>|\"#[]{}=%;",
            mainPage.GetDialogText(), "Failed! Actual message: {0}", mainPage.GetDialogText());

            //8 Close Warning Message box
            mainPage.ConfirmDialog("OK");

            //9 Click Add New link
            //10 Enter value into Display Name field with special character is @            
            chart = new Chart("Logigear@", "Name", null);
            addPanelPopup.AddChart(chart);

            //VP The new panel is created
            Assert.AreEqual(true, panelPage.IsLinkExist("Logigear@"), "Page Logigear@ cannot be created");

            panelPage.DeletePanels("All");
        }

        /// <summary>
        /// Verify that "Category", "Series" and "Caption" field are enabled and disabled correctly corresponding to each type of the "Chart Type"
        /// </summary>
        /// <author>Tu Nguyen</author>
        [TestMethod]
        public void DA_PANEL_TC037()
        {
            Console.WriteLine("DA_PANEL_TC037 - Verify that \"Category\", \"Series\" and \"Caption\" field are enabled and disabled correctly corresponding to each type of the \"Chart Type\"");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' button
            //6. Enter Page Name
            //7. Click 'OK' button
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button below 'main_hung' button        
            //9. Click 'Create new panel' button
            mainPage.OpenNewPanelPopUp(page.PageName);

            //10. Click 'Chart Type' drop-down menu
            //11. Select 'Pie' Chart Type
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);
            addPanelPopup.SelectChartType("Pie");

            //VP: Check that 'Category' and 'Caption' are disabled, 'Series' is enabled
            string actualCategoryStatus = addPanelPopup.GetCategoryStatus();
            string actualCategoryCaptionStatus = addPanelPopup.GetCategoryCaptionStatus();
            string actualSeriesStatus = addPanelPopup.GetSeriesStatus();
            string actualSeriesCaptionStatus = addPanelPopup.GetSeriesCaptionStatus();

            Assert.AreEqual("disabled", actualCategoryStatus, "Category combobox is " + actualCategoryStatus);
            Assert.AreEqual("disabled", actualCategoryCaptionStatus, "Category Caption textbox is " + actualCategoryCaptionStatus);
            Assert.AreEqual("enabled", actualSeriesStatus, "Series combobox is " + actualSeriesStatus);
            Assert.AreEqual("disabled", actualSeriesCaptionStatus, "Series Caption textbox is " + actualSeriesCaptionStatus);

            //12. Click 'Chart Type' drop-down menu
            //13. Select 'Single Bar' Chart Type
            addPanelPopup.SelectChartType("Single Bar");

            //VP: Check that 'Category' is disabled, 'Series' and 'Caption' are enabled
            string actualCategoryStatus1 = addPanelPopup.GetCategoryStatus();
            string actualCategoryCaptionStatus1 = addPanelPopup.GetCategoryCaptionStatus();
            string actualSeriesStatus1 = addPanelPopup.GetSeriesStatus();
            string actualSeriesCaptionStatus1 = addPanelPopup.GetSeriesCaptionStatus();

            Assert.AreEqual("disabled", actualCategoryStatus1, "Category combobox is " + actualCategoryStatus1);
            Assert.AreEqual("enabled", actualCategoryCaptionStatus1, "Category Caption textbox is " + actualCategoryCaptionStatus1);
            Assert.AreEqual("enabled", actualSeriesStatus1, "Series combobox is " + actualSeriesStatus1);
            Assert.AreEqual("enabled", actualSeriesCaptionStatus1, "Series Caption textbox is " + actualSeriesCaptionStatus1);

            //14. Click 'Chart Type' drop-down menu
            //15. Select 'Stacked Bar' Chart Type
            addPanelPopup.SelectChartType("Stacked Bar");

            //VP: Check that 'Category' ,'Series' and 'Caption' are enabled
            string actualCategoryStatus2 = addPanelPopup.GetCategoryStatus();
            string actualCategoryCaptionStatus2 = addPanelPopup.GetCategoryCaptionStatus();
            string actualSeriesStatus2 = addPanelPopup.GetSeriesStatus();
            string actualSeriesCaptionStatus2 = addPanelPopup.GetSeriesCaptionStatus();

            Assert.AreEqual("enabled", actualCategoryStatus2, "Category combobox is " + actualCategoryStatus2);
            Assert.AreEqual("enabled", actualCategoryCaptionStatus2, "Category Caption textbox is " + actualCategoryCaptionStatus2);
            Assert.AreEqual("enabled", actualSeriesStatus2, "Series combobox is " + actualSeriesStatus2);
            Assert.AreEqual("enabled", actualSeriesCaptionStatus2, "Series Caption textbox is " + actualSeriesCaptionStatus2);

            //16. Click 'Chart Type' drop-down menu
            //17. Select 'Group Bar' Chart Type
            addPanelPopup.SelectChartType("Group Bar");

            //VP: Check that 'Category' ,'Series' and 'Caption' are enabled
            string actualCategoryStatus3 = addPanelPopup.GetCategoryStatus();
            string actualCategoryCaptionStatus3 = addPanelPopup.GetCategoryCaptionStatus();
            string actualSeriesStatus3 = addPanelPopup.GetSeriesStatus();
            string actualSeriesCaptionStatus3 = addPanelPopup.GetSeriesCaptionStatus();

            Assert.AreEqual("enabled", actualCategoryStatus3, "Category combobox is " + actualCategoryStatus3);
            Assert.AreEqual("enabled", actualCategoryCaptionStatus3, "Category Caption textbox is " + actualCategoryCaptionStatus3);
            Assert.AreEqual("enabled", actualSeriesStatus3, "Series combobox is " + actualSeriesStatus3);
            Assert.AreEqual("enabled", actualSeriesCaptionStatus3, "Series Caption textbox is " + actualSeriesCaptionStatus3);

            //18. Click 'Chart Type' drop-down menu
            //19. Select 'Line' Chart Type
            addPanelPopup.SelectChartType("Line");

            //VP: Check that 'Category' ,'Series' and 'Caption' are enabled
            string actualCategoryStatus4 = addPanelPopup.GetCategoryStatus();
            string actualCategoryCaptionStatus4 = addPanelPopup.GetCategoryCaptionStatus();
            string actualSeriesStatus4 = addPanelPopup.GetSeriesStatus();
            string actualSeriesCaptionStatus4 = addPanelPopup.GetSeriesCaptionStatus();

            Assert.AreEqual("enabled", actualCategoryStatus4, "Category combobox is " + actualCategoryStatus4);
            Assert.AreEqual("enabled", actualCategoryCaptionStatus4, "Category Caption textbox is " + actualCategoryCaptionStatus4);
            Assert.AreEqual("enabled", actualSeriesStatus4, "Series combobox is " + actualSeriesStatus4);
            Assert.AreEqual("enabled", actualSeriesCaptionStatus4, "Series Caption textbox is " + actualSeriesCaptionStatus4);

            //Post-Condition
            addPanelPopup.ClosePanelDialog("Cancel");
            mainPage.DeletePage(page.PageName);
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that all "Data Labels" check boxes are enabled and disabled correctly corresponding to each type of "Chart Type"
        /// </summary>
        /// <author>Tu Nguyen</author>
        [TestMethod]
        public void DA_PANEL_TC040()
        {
            Console.WriteLine("DA_PANEL_TC040 - Verify that all \"Data Labels\" check boxes are enabled and disabled correctly corresponding to each type of \"Chart Type\"");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' button
            //6. Enter Page Name
            //7. Click 'OK' button
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button below 'main_hung' button        
            //9. Click 'Create new panel' button
            mainPage.OpenNewPanelPopUp(page.PageName);

            //10. Click 'Chart Type' drop-down menu
            //11. Select 'Pie' Chart Type
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);
            addPanelPopup.SelectChartType("Pie");

            //VP: Check that 'Categories' checkbox is disabled, 'Series' checkbox, 'Value' checkbox and 'Percentage' checkbox are enabled
            string actualCategoriesChbStatus = addPanelPopup.GetCategoriesCheckBoxStatus();
            string actualSeriesChbStatus = addPanelPopup.GetSeriesCheckBoxStatus();
            string actualValueChbStatus = addPanelPopup.GetValueCheckBoxStatus();
            string actualPercentageChbStatus = addPanelPopup.GetPercentageCheckBoxStatus();

            Assert.AreEqual("disabled", actualCategoriesChbStatus, "Categories checkbox is " + actualCategoriesChbStatus);
            Assert.AreEqual("enabled", actualSeriesChbStatus, "Series checkbox is " + actualSeriesChbStatus);
            Assert.AreEqual("enabled", actualValueChbStatus, "Value checkbox is " + actualValueChbStatus);
            Assert.AreEqual("enabled", actualPercentageChbStatus, "Percentage checkbox is " + actualPercentageChbStatus);

            //13. Click 'Chart Type' drop-down menu
            //14. Select 'Single Bar' Chart Type
            addPanelPopup.SelectChartType("Single Bar");

            //VP: Check that 'Categories' checkbox is disabled, 'Series' checkbox, 'Value' checkbox and 'Percentage' checkbox are enabled
            string actualCategoriesChbStatus1 = addPanelPopup.GetCategoriesCheckBoxStatus();
            string actualSeriesChbStatus1 = addPanelPopup.GetSeriesCheckBoxStatus();
            string actualValueChbStatus1 = addPanelPopup.GetValueCheckBoxStatus();
            string actualPercentageChbStatus1 = addPanelPopup.GetPercentageCheckBoxStatus();

            Assert.AreEqual("disabled", actualCategoriesChbStatus1, "Categories checkbox is " + actualCategoriesChbStatus1);
            Assert.AreEqual("enabled", actualSeriesChbStatus1, "Series checkbox is " + actualSeriesChbStatus1);
            Assert.AreEqual("enabled", actualValueChbStatus1, "Value checkbox is " + actualValueChbStatus1);
            Assert.AreEqual("enabled", actualPercentageChbStatus1, "Percentage checkbox is " + actualPercentageChbStatus1);

            //15. Click 'Chart Type' drop-down menu
            //16. Select 'Stacked Bar' Chart Type
            addPanelPopup.SelectChartType("Stacked Bar");

            //VP: Check that 'Categories' checkbox, 'Series' checkbox, 'Value' checkbox and 'Percentage' checkbox are enabled
            string actualCategoriesChbStatus2 = addPanelPopup.GetCategoriesCheckBoxStatus();
            string actualSeriesChbStatus2 = addPanelPopup.GetSeriesCheckBoxStatus();
            string actualValueChbStatus2 = addPanelPopup.GetValueCheckBoxStatus();
            string actualPercentageChbStatus2 = addPanelPopup.GetPercentageCheckBoxStatus();

            Assert.AreEqual("enabled", actualCategoriesChbStatus2, "Categories checkbox is " + actualCategoriesChbStatus2);
            Assert.AreEqual("enabled", actualSeriesChbStatus2, "Series checkbox is " + actualSeriesChbStatus2);
            Assert.AreEqual("enabled", actualValueChbStatus2, "Value checkbox is " + actualValueChbStatus2);
            Assert.AreEqual("enabled", actualPercentageChbStatus2, "Percentage checkbox is " + actualPercentageChbStatus2);

            //17. Click 'Chart Type' drop-down menu
            //18. Select 'Group Bar' Chart Type
            addPanelPopup.SelectChartType("Group Bar");

            //VP: Check that 'Categories' checkbox, 'Series' checkbox, 'Value' checkbox and 'Percentage' checkbox are enabled
            string actualCategoriesChbStatus3 = addPanelPopup.GetCategoriesCheckBoxStatus();
            string actualSeriesChbStatus3 = addPanelPopup.GetSeriesCheckBoxStatus();
            string actualValueChbStatus3 = addPanelPopup.GetValueCheckBoxStatus();
            string actualPercentageChbStatus3 = addPanelPopup.GetPercentageCheckBoxStatus();

            Assert.AreEqual("enabled", actualCategoriesChbStatus3, "Categories checkbox is " + actualCategoriesChbStatus3);
            Assert.AreEqual("enabled", actualSeriesChbStatus3, "Series checkbox is " + actualSeriesChbStatus3);
            Assert.AreEqual("enabled", actualValueChbStatus3, "Value checkbox is " + actualValueChbStatus3);
            Assert.AreEqual("enabled", actualPercentageChbStatus3, "Percentage checkbox is " + actualPercentageChbStatus3);

            //19. Click 'Chart Type' drop-down menu
            //20. Select 'Line' Chart Type
            addPanelPopup.SelectChartType("Line");

            //VP: Check that 'Categories' checkbox, 'Series' checkbox, 'Value' checkbox and 'Percentage' checkbox are disabled
            string actualCategoriesChbStatus4 = addPanelPopup.GetCategoriesCheckBoxStatus();
            string actualSeriesChbStatus4 = addPanelPopup.GetSeriesCheckBoxStatus();
            string actualValueChbStatus4 = addPanelPopup.GetValueCheckBoxStatus();
            string actualPercentageChbStatus4 = addPanelPopup.GetPercentageCheckBoxStatus();

            Assert.AreEqual("disabled", actualCategoriesChbStatus4, "Categories checkbox is " + actualCategoriesChbStatus4);
            Assert.AreEqual("disabled", actualSeriesChbStatus4, "Series checkbox is " + actualSeriesChbStatus4);
            Assert.AreEqual("disabled", actualPercentageChbStatus4, "Percentage checkbox is " + actualPercentageChbStatus4);
            Assert.AreEqual("disabled", actualValueChbStatus4, "Value checkbox is " + actualValueChbStatus4);

        }

        /// <summary>
        /// Verify that all settings within "Add New Panel" and "Edit Panel" form stay unchanged when user switches between "2D" and "3D" radio buttons
        /// </summary>
        /// <author>Tu Nguyen</author>
        [TestMethod]
        public void DA_PANEL_TC038()
        {
            Console.WriteLine("DA_PANEL_TC038 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"2D\" and \"3D\" radio buttons");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' button
            //6. Enter Page Name
            //7. Click 'OK' button
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button below 'main_hung' button        
            //9. Click 'Create new panel' button
            PanelsPage panelPage = new PanelsPage();
            mainPage.OpenNewPanelPopUp(page.PageName);
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);

            //10. Click 'Chart Type' drop-down menu
            //11. Select 'Stacked Bar' Chart Type
            //12. Select 'Data Profile' drop-down menu
            //13. Enter 'Display Name' and 'Chart Title'
            //14. Select 'Show Title' checkbox
            //15. Select 'Legends' radio button
            //16. Select 'Style' radio button
            string displayName = CommonAction.GeneratePanelName();
            string chartTitle = "Tu_" + CommonAction.GeneratePanelName();
            addPanelPopup.FillPanelData(null, "Test Case Execution", displayName, chartTitle, "on", "Stacked Bar", "3D", null, null, null, true, "Top");

            //VP: Settings of 'Chart Type', 'Data Profile', 'Display Name', 'Chart Title', 'Show Title' and 'Legends' stay unchanged
            string currentChartType = addPanelPopup.CbbChartType.Value;
            string currentProfile = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle = addPanelPopup.TxtChartTitle.Value;
            string currentShowTitle = addPanelPopup.ChbShowTitle.Value;
            string currentLegend = addPanelPopup.GetLegend();

            Assert.AreEqual("Stacked Bar", currentChartType, "Current chart type is " + currentChartType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Data Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Display name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Show Title is " + currentShowTitle);
            Assert.AreEqual("Top", currentLegend, "Legend is " + currentLegend);

            //17. Select 'Style' radio button
            addPanelPopup.SelectStyle("2D");

            //VP: Settings of 'Chart Type', 'Data Profile', 'Display Name', 'Chart Title', 'Show Title' and 'Legends' stay unchanged
            Assert.AreEqual("Stacked Bar", currentChartType, "Current chart type is " + currentChartType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Data Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Display name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Show Title is " + currentShowTitle);
            Assert.AreEqual("Top", currentLegend, "Legend is " + currentLegend);

            //Close Add Panel Page
            //Open Add New Panel Page
            addPanelPopup.ClosePanelDialog("Cancel");
            mainPage.OpenNewPanelPopUp(page.PageName);

            //19. Select a page in drop-down menu
            //20. Enter path of Folder
            //21. Click OK button
            Chart chart = new Chart("Test Case Execution", displayName, page.PageName, 400, null, chartTitle, "Stacked Bar", "Name", null, "Location", null, null, null, "2D", false);
            addPanelPopup.AddChart(chart);

            //22. Click 'Edit Panel' button of panel 'hung_panel'
            mainPage.OpenEditPanelPopup();

            //23. Select 'Style' radio button
            addPanelPopup.SelectStyle("3D");

            //VP: Check that settings of 'Chart Type', 'Data Profile', 'Display Name', 'Chart Title', 'Show Title' and 'Legends' stay unchanged.
            string currentChartType1 = addPanelPopup.CbbChartType.Value;
            string currentProfile1 = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName1 = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle1 = addPanelPopup.TxtChartTitle.Value;
            string currentShowTitle1 = addPanelPopup.ChbShowTitle.Value;
            string currentLegend1 = addPanelPopup.GetLegend();

            Assert.AreEqual("Stacked Bar", currentChartType1, "Current chart type is " + currentChartType1);
            Assert.AreEqual("Test Case Execution", currentProfile1, "Data Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Display name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Chart Title is " + currentChartTitle1);
            Assert.AreEqual("on", currentShowTitle1, "Show Title is " + currentShowTitle1);
            Assert.AreEqual("Top", currentLegend1, "Legend is " + currentLegend1);

            //24. Select 'Style' radio button
            addPanelPopup.SelectStyle("2D");

            //VP: Check that settings of 'Chart Type', 'Data Profile', 'Display Name', 'Chart Title', 'Show Title' and 'Legends' stay unchanged.
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current chart type is " + currentChartType1);
            Assert.AreEqual("Test Case Execution", currentProfile1, "Data Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Display name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Chart Title is " + currentChartTitle1);
            Assert.AreEqual("on", currentShowTitle1, "Show Title is " + currentShowTitle1);
            Assert.AreEqual("Top", currentLegend1, "Legend is " + currentLegend1);

            //Post Condition            
            panelPage.DeletePanels("All");
            addPanelPopup.ClosePanelDialog("Cancel");
            mainPage.DeletePage(page.PageName);

        }

        /// <summary>
        /// Verify that all settings within "Add New Panel" and "Edit Panel" form stay unchanged when user switches between "Legends" radio buttons
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC039()
        {
            Console.WriteLine("DA_PANEL_TC039 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"Legends\" radio buttons");

            //Set Variables
            string displayName = CommonAction.GeneratePanelName();
            string chartTitle = "Tu" + CommonAction.GeneratePanelName();

            //1 Navigate to Dashboard login page
            //2 Select specific repository
            //3 Enter valid username and password
            //4 Click on Login button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3 Click Administer/Panels link            
            //5 Click Add New link

            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.OpenAddNewPanelPopupFromLink();
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);
            addPanelPopup.FillPanelData("Chart", "Test Case Execution", displayName, chartTitle, "on", "Stacked Bar", "2D", "Name", "Location", null, true, null);
            string currentType = addPanelPopup.GetPanelType();
            string currentProfile = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle = addPanelPopup.TxtChartTitle.Value;
            string currentShowTitle = addPanelPopup.ChbShowTitle.Value;
            string currentChartType = addPanelPopup.CbbChartType.Value;
            string currentStyle = addPanelPopup.GetStyle();
            string currentCategory = addPanelPopup.CbbCategory.Value;
            string currentSeries = addPanelPopup.CbbSeries.Value;

            //6. Click None radio button for Legend
            addPanelPopup.SelectLegend("None");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //7. Click Top radio button for Legend
            addPanelPopup.SelectLegend("Top");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //8. Click Right radio button for Legend
            addPanelPopup.SelectLegend("Right");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //9. Click Bottom radio button for Legend
            addPanelPopup.SelectLegend("Bottom");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //10. Click Left radio button for Legend
            addPanelPopup.SelectLegend("Left");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //11. Create a new panel
            addPanelPopup.ClosePanelDialog("Cancel");
            panelPage.OpenAddNewPanelPopupFromLink();
            Chart chart = new Chart("Test Case Execution", displayName, null, 400, null, chartTitle, "Stacked Bar", "Name", null, "Location", null, null, null, "2D", false);
            addPanelPopup.AddChart(chart);

            //12. Click Edit Panel link
            panelPage.OpenEditPanelPopup(displayName);

            string currentProfile1 = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName1 = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle1 = addPanelPopup.TxtChartTitle.Value;
            string currentChartType1 = addPanelPopup.CbbChartType.Value;
            string currentStyle1 = addPanelPopup.GetStyle();
            string currentCategory1 = addPanelPopup.CbbCategory.Value;
            string currentSeries1 = addPanelPopup.CbbSeries.Value;

            //13. Click None radio button for Legend
            addPanelPopup.SelectLegend("None");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //14. Click Top radio button for Legend
            addPanelPopup.SelectLegend("Top");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //15. Click Right radio button for Legend
            addPanelPopup.SelectLegend("Right");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //16. Click Bottom radio button for Legend
            addPanelPopup.SelectLegend("Bottom");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //17. Click Left radio button for Legend
            addPanelPopup.SelectLegend("Left");

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //Post-Condition
            addPanelPopup.ClosePanelDialog("Cancel");
            panelPage.DeletePanels("All");
            
        }

        /// <summary>
        /// Verify that all settings within "Add New Panel" and "Edit Panel" form stay unchanged when user switches between "Data Labels" check boxes buttons.
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC041()
        {
            Console.WriteLine("DA_PANEL_TC041 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"Data Labels\" check boxes buttons");

            //Set Variables
            string displayName = CommonAction.GeneratePanelName();
            string chartTitle = "Tu" + CommonAction.GeneratePanelName();

            //1 Navigate to Dashboard login page
            //2 Select specific repository
            //3 Enter valid username and password
            //4 Click on Login button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3 Click Administer/Panels link            
            //5 Click Add New link
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.OpenAddNewPanelPopupFromLink();
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);
            addPanelPopup.FillPanelData("Chart", "Test Case Execution", displayName, chartTitle, "on", "Stacked Bar", "2D", "Name", "Location", null, true, null);
            string currentType = addPanelPopup.GetPanelType();
            string currentProfile = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle = addPanelPopup.TxtChartTitle.Value;
            string currentShowTitle = addPanelPopup.ChbShowTitle.Value;
            string currentChartType = addPanelPopup.CbbChartType.Value;
            string currentStyle = addPanelPopup.GetStyle();
            string currentCategory = addPanelPopup.CbbCategory.Value;
            string currentSeries = addPanelPopup.CbbSeries.Value;

            //6. Check Series checkbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Series", true, null);

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //7. Uncheck Series checkbox
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Series", false, null);

            //8. Check Value checkbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Value", true, null);

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //9. Uncheck Value checkbox
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Value", false, null);

            //10. Check Percentage checbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Percentage", true, null);

            //VP: All settings are unchange in Add New Panel dialog
            Assert.AreEqual("Chart", currentType, "Current Type is " + currentType);
            Assert.AreEqual("Test Case Execution", currentProfile, "Current Profile is " + currentProfile);
            Assert.AreEqual(displayName, currentName, "Current Display Name is " + currentName);
            Assert.AreEqual(chartTitle, currentChartTitle, "Current Chart Title is " + currentChartTitle);
            Assert.AreEqual("on", currentShowTitle, "Current Show Title is " + currentShowTitle);
            Assert.AreEqual("Stacked Bar", currentChartType, "Current Chart Type is " + currentChartType);
            Assert.AreEqual("2D", currentStyle, "Current Style is " + currentStyle);
            Assert.AreEqual("name", currentCategory, "Current Category is " + currentCategory);
            Assert.AreEqual("location", currentSeries, "Current Series is " + currentSeries);

            //11. Uncheck Percentage checkbox
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Percentage", false, null);

            //Close current panel
            addPanelPopup.ClosePanelDialog("Cancel");

            //12. Create a new panel
            //13. Click Edit Panel link
            panelPage.OpenAddNewPanelPopupFromLink();
            Chart chart = new Chart("Test Case Execution", displayName, null, 400, null, chartTitle, "Stacked Bar", "Name", null, "Location", null, null, null, "2D", false);
            addPanelPopup.AddChart(chart);
            panelPage.OpenEditPanelPopup(displayName);

            string currentProfile1 = addPanelPopup.CbbDataProfile.GetSelectedText();
            string currentName1 = addPanelPopup.TxtDisplayName.Value;
            string currentChartTitle1 = addPanelPopup.TxtChartTitle.Value;
            string currentChartType1 = addPanelPopup.CbbChartType.Value;
            string currentStyle1 = addPanelPopup.GetStyle();
            string currentCategory1 = addPanelPopup.CbbCategory.Value;
            string currentSeries1 = addPanelPopup.CbbSeries.Value;

            //14. Check Series checkbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Series", true, null);

            //VP: All settings are unchange in Edit New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //15. Uncheck Series checkbox
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Series", false, null);

            //16. Check Value checkbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Value", true, null);

            //VP: All settings are unchange in Edit New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //17. Uncheck Value checkbox
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Value", false, null);

            //18. Check Percentage checbox for Data Labels
            addPanelPopup.FillPanelData(null, null, null, null, null, null, null, null, null, "Percentage", true, null);

            //VP: All settings are unchange in Edit New Panel dialog
            Assert.AreEqual("Test Case Execution", currentProfile1, "Current Profile is " + currentProfile1);
            Assert.AreEqual(displayName, currentName1, "Current Display Name is " + currentName1);
            Assert.AreEqual(chartTitle, currentChartTitle1, "Current Chart Title is " + currentChartTitle1);
            Assert.AreEqual("Stacked Bar", currentChartType1, "Current Chart Type is " + currentChartType1);
            Assert.AreEqual("2D", currentStyle1, "Current Style is " + currentStyle1);
            Assert.AreEqual("name", currentCategory1, "Current Category is " + currentCategory1);
            Assert.AreEqual("location", currentSeries1, "Current Series is " + currentSeries1);

            //Post-Condition
            addPanelPopup.ClosePanelDialog("Cancel");
            panelPage.DeletePanels("All");

        }

        /// <summary>
        /// Verify that user is unable to edit  "Height *" field to anything apart from integer number with in 300-800 range.
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC052()
        {
            Console.WriteLine("DA_PANEL_TC052 - Verify that user is unable to edit  \"Height *\" field to anything apart from integer number with in 300-800 range");

            //Set Variables
            string displayName = CommonAction.GeneratePanelName();
            string chartTitle = "Tu" + CommonAction.GeneratePanelName();

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' button
            //6. Enter Page Name
            //7. Click 'OK' button
            string pageName = CommonAction.GenrateRandomString(Constants.lenghtRandomString);
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button below 'main_hung' button        
            //9. Click 'Create new panel' button
            PanelsPage panelPage = new PanelsPage(_webDriver);
            mainPage.OpenNewPanelPopUp(page.PageName);
            AddNewPanelPage addPanelPopup = new AddNewPanelPage(_webDriver);

            //10. Enter all required fields on Add New Panel page
            addPanelPopup.FillPanelData("Chart", "Test Case Execution", displayName, chartTitle, "on", "Stacked Bar", "2D", "Name", "Location", null, true, null);

            //11. Click Ok button
            addPanelPopup.ClosePanelDialog("OK");

            //12. Enter invalid height into Height field
            //13. Click Ok button
            string actualMessage = addPanelPopup.SettingPanelWithExpectedError(page.PageName, 200, null);

            //VP: There is message "Panel Height must be greater than or equal to 300 and lower than or equal to 800"
            string expectedMessage = "Panel height must be greater than or equal to 300 and less than or equal to 800.";

            Assert.AreEqual(expectedMessage, actualMessage, "Actual Message is " + actualMessage);
            panelPage.ConfirmDialog("OK");

            //14. Enter valid height into Height field
            //15. Click Ok button
            addPanelPopup.SettingPanel(page.PageName, 400, null);
            mainPage.DeletePage(page.PageName);

            //VP:User is able to edit Height field to anything apart from integer number with in 300-800 range
            mainPage.OpenPanelsPage();
            bool actual = panelPage.IsPanelExist(displayName);
            Assert.AreEqual(true, actual, "Actual panel not exist");

            //Post-condition
            panelPage.DeletePanels("All");

        }
        /// <summary>
        /// Verify that user is not allowed to create panel with duplicated "Display Name"  			
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        [TestMethod]
        public void DA_PANEL_TC032()
        {
            Console.WriteLine("DA_PANEL_TC032 - Verify that user is not allowed to create panel with duplicated \"Display Name\"");

            //Set variables
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click on Administer/Panels link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //4. Click on Add new link
            //5. Enter display name to "Display name" field.
            //6. Click on OK button
            panelPage.AddNewPanel(chart);

            //7. Click on Add new link again.
            //8. Enter display name same with previous display name to "display name" field. 
            //9. Click on OK button
            string actualMsg = panelPage.AddChartWithExpectedError(chart);

            //VP. Warning message: "Dupicated panel already exists. Please enter a different name" show up
            string expectedMsg = string.Format("{0} already exists. Please enter a different name.", chart.DisplayName);
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));

            //Post-condition
            panelPage.ConfirmDialog("OK");
            panelPage.CancelPanel();
            panelPage.DeletePanels("All");

        }

        /// <summary>
        /// Verify that "Data Profile" listing of "Add New Panel" and "Edit Panel" control/form are in alphabetical order			
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        [TestMethod]
        public void DA_PANEL_TC033()
        {
            Console.WriteLine("DA_PANEL_TC033 - Verify that \"Data Profile\" listing of \"Add New Panel\" and \"Edit Panel\" control/form are in alphabetical order");

            //Set variables
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click on Administer/Panels link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //4. Click on Add new link
            //VP. Verify that Data Profile list is in alphabetical order
            panelPage.OpenAddNewPanelPopupFromLink();
            Assert.IsTrue(panelPage.IsDataProfileSOrder("ASC"), "Data Profile list is not in alphabetical order");

            //5. Enter a display name to display name field
            //6. Click on OK button
            panelPage.AddNewPanel(chart, false);

            //7. Click on Edit link
            //VP. Verify that Data Profile list is in alphabetical order
            panelPage.OpenEditPanelPopup(chart.DisplayName);
            Assert.IsTrue(panelPage.IsDataProfileSOrder("ASC"), "Data Profile list is not in alphabetical order");

            //Post-condition
            panelPage.CancelPanel();
            panelPage.DeletePanels("All");

        }

        /// <summary>
        /// Verify that newly created data profiles are populated correctly under the "Data Profile" dropped down menu in  "Add New Panel" and "Edit Panel" control/form			
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        [TestMethod]
        public void DA_PANEL_TC034()
        {
            Console.WriteLine("DA_PANEL_TC034 - Verify that newly created data profiles are populated correctly under the \"Data Profile\" dropped down menu in  \"Add New Panel\" and \"Edit Panel\" control/form");

            //Set variables
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);
            string dataProfileName = CommonAction.GenerateDataProfileName();

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);


            //2. Click on Administer/Data Profiles link
            //3. Click on add new link
            //4. Enter name to Name textbox
            //5. Click on Finish button
            mainPage.GoToDataProfilePage()
                    .GoToGeneralSettingPage()
                    .SetGeneralSettingsValue(dataProfileName, null, null, "Finish");

            //6. Click on Administer/Panels link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //7. Click on add new link
            //VP. Verify that "giang - data" data profiles are populated correctly under the "Data Profile" dropped down menu.
            AddNewPanelPage addNewPanelPage = panelPage.OpenAddNewPanelPopupFromLink();
            Assert.IsTrue(addNewPanelPage.IsDataProfileExists(dataProfileName), string.Format("{0} is not populated correctly under the \"Data Profile\" dropped down menu", dataProfileName));

            //8. Enter display name to Display Name textbox
            //9. Click Ok button to create a panel
            panelPage.AddNewPanel(chart, false);

            //10. Click on edit link
            //VP. Verify that "giang - data" data profiles are populated correctly under the "Data Profile" dropped down menu.
            panelPage.OpenEditPanelPopup(chart.DisplayName);
            Assert.IsTrue(addNewPanelPage.IsDataProfileExists(dataProfileName), string.Format("{0} is not populated correctly under the \"Data Profile\" dropped down menu", dataProfileName));

            //Post-condition
            panelPage.CancelPanel();
            panelPage.DeletePanels("All");
            panelPage.GoToDataProfilePage().DeleteAllDataProfiles();
        }

        /// <summary>
        /// Verify that no special character except '@' character is allowed to be inputted into "Chart Title" field			
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        [TestMethod]
        public void DA_PANEL_TC035()
        {
            Console.WriteLine("DA_PANEL_TC035 - Verify that no special character except '@' character is allowed to be inputted into \"Chart Title\" field");

            //Set variables
            Chart chart1 = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "!#$%^&*()'", null, null, null, "Name", null, null, null, null, false);
            Chart chart2 = new Chart(null, CommonAction.GeneratePanelName() + "@", null, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click on Administer/Panels link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //4. Click on Add new link
            //5. Enter value into Display Name field
            //6. Enter value into Chart Title field with special characters except "@"
            //7. Click Ok button
            string actualMsg = panelPage.AddChartWithExpectedError(chart1);

            //VP. Message "Invalid display name. The name can't contain high ASCII characters or any of following characters: /:*?<>|"#{[]{};" is displayed
            string expectedMsg = "Invalid title name. The name cannot contain high ASCII characters or any of the following characters: /:*?<>|\"#[]{}=%;";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));

            //8. Close Warning Message box
            panelPage.ConfirmDialog("OK");
            panelPage.CancelPanel();

            //9. Click Add New link
            //10. Enter value into Display Name field
            //11. Enter value into Chart Title field with special character is @
            //VP. The new panel is created
            panelPage.AddNewPanel(chart2);
            Assert.IsTrue(panelPage.IsPanelExist(chart2.DisplayName), string.Format("{0} is not created successfully ", chart2.DisplayName));

            //Post-condition
            panelPage.DeletePanels("All");
        }

        /// <summary>
        /// Verify that all chart types ( Pie, Single Bar, Stacked Bar, Group Bar, Line ) are listed correctly under "Chart Type" dropped down menu.			
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/30/2016</date>
        [TestMethod]
        public void DA_PANEL_TC036()
        {
            Console.WriteLine("DA_PANEL_TC036 - Verify that all chart types ( Pie, Single Bar, Stacked Bar, Group Bar, Line ) are listed correctly under \"Chart Type\" dropped down menu.");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Chart chart = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "!#$%^&*()'", "Name", null, null, null, null, null, null, null, false);

            //1. Navigate to Dashboard login page
            //2. Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' link
            //6. Enter Page Name
            //7. Click 'OK' button
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button
            //9. Click 'Create new panel' button
            AddNewPanelPage addNewPanelPage = mainPage.OpenAddNewPanelPageFromButton();

            //10. Click 'Chart Type' drop-down menu
            //VP. Check that 'Chart Type' are listed 5 options: 'Pie', 'Single Bar', 'Stacked Bar', 'Group Bar' and 'Line'
            string[] chartTypeOptions = { "Pie", "Single Bar", "Stacked Bar", "Group Bar", "Line" };
            Assert.IsTrue(addNewPanelPage.IsComboboxListed(addNewPanelPage.CbbChartType, chartTypeOptions), "Chart Type combobox is not listed correcty!");

            //Post-Condition
            addNewPanelPage.BtnCancel.Click();
            mainPage.DeletePage(page.PageName);
        }

        /// <summary>
        /// DA_PANEL_TC042 - Verify that all pages are listed correctly under the \"Select page *\" dropped down menu of \"Panel Configuration\" form/ control
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/03/2016</date>
        /// 
        [TestMethod]
        public void DA_PANEL_TC042()
        {
            Console.WriteLine("DA_PANEL_TC042 - Verify that all pages are listed correctly under the \"Select page *\" dropped down menu of \"Panel Configuration\" form/ control");

            //Set variables
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page page2 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page page3 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            string[] pageList = new string[] { page1.PageName, page2.PageName, page3.PageName };

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click 'Login' button
            //5 Click 'Add Page' button
            //6 Enter Page Name
            //7 Click 'OK' button
            //8 Click 'Add Page' button
            //9 Enter Page Name
            //10 Click 'OK' button
            //11 Click 'Add Page' button
            //12 Enter Page Name
            //13 Click 'OK' button
            //14 Click 'Choose panels' button
            //15 Click on any Chart panel instance
            //16 Click 'Select Page*' drop-down menu
            LoginPage loginPage = new LoginPage(_webDriver);
            MainPage mainPage = loginPage.Open().Login(Constants.Repository, Constants.UserName, Constants.Password);
            mainPage.AddPage(page1)
                .AddPage(page2)
                .AddPage(page3);
            AddNewPanelPage configurationPopup = mainPage.OpenPanelConfigurationFromChoosePanel("Test Case Execution Results");
            configurationPopup.isComboboxContainsItems(configurationPopup.CbbSelectPage, pageList);

            //17 Check that 'Select Page*' drop-down menu contains 3 items: 'main_hung1', 'main_hung2' and 'main_hung3'
            Assert.AreEqual(true, configurationPopup.isComboboxContainsItems(configurationPopup.CbbSelectPage, pageList), "Select Page * Combobox does not contain these page name");

            //Post-Condition
            configurationPopup.BtnCancel.Click();
            mainPage.WaitForPageLoadComplete();
            mainPage.DeletePage(page1.PageName)
                .DeletePage(page2.PageName)
                .DeletePage(page3.PageName)
                .Logout();
        }
        /// <summary>
        /// DA_PANEL_TC043 - Verify that only integer number inputs from 300-800 are valid for \"Height *\" field 
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/03/2016</date>
        [TestMethod]
        public void DA_PANEL_TC043()
        {
            Console.WriteLine("DA_PANEL_TC043 - Verify that only integer number inputs from 300-800 are valid for \"Height *\" field ");
            //Set variables
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click 'Login' button
            //5 Click 'Add Page' button
            //6 Enter Page Name
            //7 Click 'OK' button
            //8 Click 'Choose panels' button
            //9 Click on any Chart panel instance
            //10 Enter integer number to 'Height *' field '299
            //11 Click OK button
            
            LoginPage loginPage = new LoginPage(_webDriver);
            MainPage mainPage = loginPage.Open().Login(Constants.Repository, Constants.UserName, Constants.Password);
            mainPage.AddPage(page1);
            AddNewPanelPage configurationPopup = mainPage.OpenPanelConfigurationFromChoosePanel("Test Case Execution Results");
            string actualMessage = configurationPopup.SettingPanelWithExpectedError(null, 299, null);

            //12 Check that error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            //VP Error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            Assert.AreEqual("Panel height must be greater than or equal to 300 and less than or equal to 800.",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");

            //13 Click OK button
            //14 Enter integer number to 'Height *' field 801
            //15 Click OK button
            //16 Check that error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            //VP Error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            actualMessage = configurationPopup.SettingPanelWithExpectedError(null, 801, null);
            Assert.AreEqual("Panel height must be greater than or equal to 300 and less than or equal to 800.",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");

            //17 Click OK button
            //18 Enter integer number to 'Height *' field -2
            //19 Click OK button
            //20 Check that error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            //VP Error message 'Panel height must be greater than or equal to 300 and lower than or equal to 800' display
            actualMessage = configurationPopup.SettingPanelWithExpectedError(null, -2, null);
            Assert.AreEqual("Panel height must be greater than or equal to 300 and less than or equal to 800.",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");

            //21 Click OK button
            //22 Enter integer number to 'Height *' field 3.1
            //23 Click OK button
            //24 Check that error message 'Panel height must be an integer number' display
            //VP Error message 'Panel height must be an integer number' display
            actualMessage = configurationPopup.SettingPanelWithExpectedError(null, 3.1, null);
            Assert.AreNotEqual("Panel height must be an integer number",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");

            //25 Click OK button
            //26 Enter integer number to 'Height *' field ABC
            //27 Click OK button
            //28 Check that error message 'Panel height must be an integer number' display
            //VP Error message 'Panel height must be an integer number' display
            actualMessage = configurationPopup.SettingPanelWithExpectedError(null, "abc", null);
            Assert.AreEqual("Panel height must be an integer number",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");
            configurationPopup.BtnCancel.Click();

            //Post Condition
            mainPage.DeletePage(page1.PageName).Logout();
        }
        /// <summary>
        /// DA_PANEL_TC044 - Verify that \"Height *\" field is not allowed to be empty
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/04/2016</date>
        [TestMethod]
        public void DA_PANEL_TC044()
        {
            Console.WriteLine("DA_PANEL_TC044 - Verify that \"Height *\" field is not allowed to be empty");
            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click 'Login' button
            //5 Click 'Add Page' button
            //6 Enter Page Name
            //7 Click 'OK' button
            //8 Click 'Choose panels' button
            //9 Click on any Chart panel instance
            //10 Leave 'Height *' field empty
            //11 Click OK button            
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            LoginPage loginPage = new LoginPage(_webDriver);
            MainPage mainPage = loginPage.Open().Login(Constants.Repository, Constants.UserName, Constants.Password);
            mainPage.AddPage(page1);
            AddNewPanelPage configurationPopup = mainPage.OpenPanelConfigurationFromChoosePanel("Test Case Execution Results");
            string actualMessage = configurationPopup.SettingPanelWithExpectedError(null, "", null);

            //12 Check that 'Panel height is required field' message display
            //VP 'Panel height is required field' message display
            Assert.AreEqual("Panel height is a required field.",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");
            configurationPopup.BtnCancel.Click();

            //Post Condition
            mainPage.DeletePage(page1.PageName).Logout();
        }
        /// <summary>
        /// DA_PANEL_TC045 - Verify that \"Folder\" field is not allowed to be empty
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/04/2016</date>
        [TestMethod]
        public void DA_PANEL_TC045()
        {
            Console.WriteLine("DA_PANEL_TC045 - Verify that \"Folder\" field is not allowed to be empty");
            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Create a new page
            //4 Click Choose Panel button
            //5 Click Create New Panel button
            //6 Enter all required fields on Add New Panel page
            //7 Click Ok button
            //8 Leave empty on Folder field
            //9 Click Ok button on Panel Configuration dialog          
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            LoginPage loginPage = new LoginPage(_webDriver);
            MainPage mainPage = loginPage.Open().Login(Constants.Repository, Constants.UserName, Constants.Password);
            mainPage.AddPage(page1);
            AddNewPanelPage configurationPopup = mainPage.OpenPanelConfigurationFromChoosePanel("Test Case Execution Results");
            string actualMessage = configurationPopup.SettingPanelWithExpectedError(null, null, " ");

            //10 Observe the current page
            //VP There is message "Panel folder is incorrect"
            Assert.AreEqual("Panel folder is incorrect",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            configurationPopup.ConfirmDialog("OK");
            configurationPopup.BtnCancel.Click();

            //Post Condition
            mainPage.DeletePage(page1.PageName).Logout();
        }
        /// <summary>
        /// DA_PANEL_TC046 - Verify that only valid folder path of corresponding item type ( e.g. Actions, Test Modules) are allowed to be entered into "Folder" field
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/04/2016</date>
        [TestMethod]
        public void DA_PANEL_TC046()
        {
            Console.WriteLine("DA_PANEL_TC046 - Verify that only valid folder path of corresponding item type ( e.g. Actions, Test Modules) are allowed to be entered into \"Folder\" field");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Create a new page
            //4 Click Choose Panel button
            //5 Click Create New Panel button
            //6 Enter all required fields on Add New Panel page
            //7 Click Ok button
            //8 Enter invalid folder path
            //9 Click Ok button on Panel Configuration dialog           
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", page1.PageName);
            chart.Folder = "abc";
            LoginPage loginPage = new LoginPage(_webDriver);
            MainPage mainPage = loginPage.Open().Login(Constants.Repository, Constants.UserName, Constants.Password);
            AddNewPanelPage addPanelPopup = mainPage.AddPage(page1).GoToPage(page1.PageName).OpenAddNewPanelPage();
            //AddNewPanelPage addPanelPopup = panelPage.OpenAddNewPanelPage();
            string errorMessage = addPanelPopup.AddChartWithExpectedError(chart);
            //VP Observe the current page.There is message "Panel folder is incorrect"
            Assert.AreEqual("Panel folder is incorrect",
               errorMessage,
               string.Format("Failed! Actual message is: {0}", errorMessage));            

            //11 Enter valid folder path
            //12 Click Ok button on Panel Configuration dialog
            chart.Folder = "/Car Rental/Tests";
            addPanelPopup.SettingPanel(chart.PageName, chart.Height, chart.Folder);

            //VP Observe the current page -The new panel is created
            Assert.AreEqual(true, mainPage.IsDivExist(chart.DisplayName), "Panel cannot be created");

            //post condition
            mainPage.DeletePage(page1.PageName).OpenPanelsPage().DeletePanels(chart.DisplayName).Logout();
        }

        /// <summary>
        /// Verify that user is able to navigate properly to folders with "Select Folder" form
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC047()
        {
            Console.WriteLine("DA_PANEL_TC047 - Verify that user is able to navigate properly to folders with \"Select Folder\" form");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            string panelName = CommonAction.GeneratePanelName();

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            mainPage.AddPage(page);

            //4. Click Choose Panel button
            //5. Click Create New Panel button
            //6. Enter all required fields on Add New Panel page
            //7. Click Ok button
            AddNewPanelPage addNewPanelPage = mainPage.OpenAddNewPanelPage().FillPanelData(null, "Test Case Execution", panelName, "test", "on", "Pie", "3D", null, "Name", null, true, "Top");
            addNewPanelPage.BtnOk.Click();

            //8. Click Select Folder button on Panel Configuration dialog
            //9. Choose folder name in Folder Form
            //10. Click Ok button on Select Folder form
            addNewPanelPage.SelectFolder("/Car Rental/Tests");

            //VP. User is able to select properly folder with Select Folder form
            Assert.AreEqual("/Car Rental/Tests", addNewPanelPage.TxtFolder.GetAttribute("value"), "Select folder is unsuccessfully");

            //Post-Condition
            addNewPanelPage.BtnOKConfigurationPanel.Click();
            mainPage.OpenPanelsPage().DeletePanels("All");
            mainPage.DeletePage(page.PageName);

        }

        /// <summary>
        /// Verify that population of corresponding item type ( e.g. Actions, Test Modules) folders is correct in "Select Folder form
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC048()
        {
            Console.WriteLine("DA_PANEL_TC048 - Verify that population of corresponding item type ( e.g. Actions, Test Modules) folders is correct in \"Select Folder form\"");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            mainPage.AddPage(page);

            //4. Click Choose Panel button
            //5. Click Create New Panel button
            //6. Enter all required fields on Add New Panel page
            //7. Click Ok button
            AddNewPanelPage addPanelPage = mainPage.OpenAddNewPanelPageFromButton().FillPanelData(null, "Test Case Execution", "Tu_Panel", "Tu_Title", "on", "Stacked Bar", "3D", null, null, null, true, "Top");
            addPanelPage.BtnOk.Click();

            //8. Click Select Folder button on Panel Configuration dialog
            //VP. Population of corresponding item type ( e.g. Actions, Test Modules) folders is correct in "Select Folder form
            Assert.IsTrue(addPanelPage.IsSelectFolderPopulation("/Car Rental/Tests"), string.Format("Folder field displays incorrently: {0}", addPanelPage.TxtFolder.Text));
        }

        /// <summary>
        /// Verify that all folder paths of corresponding item type ( e.g. Actions, Test Modules) are correct in "Select Folder" form 
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC049()
        {
            Console.WriteLine("DA_PANEL_TC049 - Verify that all folder paths of corresponding item type ( e.g. Actions, Test Modules) are correct in \"Select Folder\" form ");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            string panelName = CommonAction.GeneratePanelName();

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            mainPage.AddPage(page);

            //4. Click Choose Panel button
            //5. Click Create New Panel button
            //6. Enter all required fields on Add New Panel page
            //7. Click Ok button
            mainPage.GoToPage(page.PageName);
            AddNewPanelPage addPanelPage = mainPage.OpenAddNewPanelPageFromButton().FillPanelData(null, "Test Case Execution", panelName, "test", "on", "Pie", "3D", null, "Name", null, true, "Top");
            addPanelPage.BtnOk.Click();

            //8. Click Select Folder button on Panel Configuration dialog
            //9. Choose folder name in Folder Form
            //10. Click Ok button on Select Folder form
            //VP. Folder path is displayed correctly after selecting folder in Select Folder form
            addPanelPage.SelectFolder("/Car Rental/Tests");
            addPanelPage.BtnOKConfigurationPanel.Click();
            Assert.AreEqual("/Car Rental/Tests", addPanelPage.TxtFolder.GetAttribute("value"), string.Format("Folder field displays incorrently: {0}", addPanelPage.TxtFolder.GetAttribute("value")));

        }

        /// <summary>
        /// Verify that user is able to successfully edit "Display Name" of any Panel providing that the name is not duplicated with existing Panels' name
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC050()
        {
            Console.WriteLine("DA_PANEL_TC050 - Verify that user is able to successfully edit \"Display Name\" of any Panel providing that the name is not duplicated with existing Panels' name");

            //Set variables
            Chart chart = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "test", null, null, null, "Name", null, null, null, null, false);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click Administer link
            //4. Click Panel link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //5. Click Add New link
            //6. Enter a valid name into Display Name field
            panelPage.AddNewPanel(chart);

            //VP. The new panel is created successfully
            panelPage.IsPanelExist(chart.DisplayName);

            //Post-Condition
            panelPage.DeletePanels("All");

        }

        /// <summary>
        /// Verify that user is unable to change "Display Name" of any Panel if there is special character except '@' inputted
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC051()
        {
            Console.WriteLine("DA_PANEL_TC051 - Verify that all pages are listed correctly under the \"Select page *\" dropped down menu of \"Panel Configuration\" form/ control");

            //AppDomainSetup variable
            Chart chart1 = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);
            Chart chart2 = new Chart(null, "/:*?<>|\"#", null, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click Administer link
            //4. Click Panel link
            PanelsPage panelPage = mainPage.OpenPanelsPage();

            //5. Click Add New link
            //6. Create a new panel
            panelPage.AddNewPanel(chart1);

            //7. Click Edit link
            //8. Edit panel name with special characters
            //9. Click Ok button
            panelPage.OpenEditPanelPopup(chart1.DisplayName);
            panelPage.EditChartPanels(chart2);

            //VP. Message "Invalid display name. The name can't contain high ASCII characters or any of following characters: /:*?<>|"#{[]{};" is displayed
            //10. Close warning message box
            string actualMsg = panelPage.GetDialogText();
            string expectedMsg = "Invalid display name. The name cannot contain high ASCII characters or any of the following characters: /:*?<>|\"#[]{}=%;";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
            panelPage.ConfirmDialog("OK");
            panelPage.CancelPanel();

            //11. Click Edit link
            //12. Edit panel name with special character is @
            //13. Click Ok button
            //VP. User is able to edit panel name with special characters is @
            panelPage.OpenEditPanelPopup(chart1.DisplayName);
            chart1.DisplayName = chart1.DisplayName + "@";
            panelPage.EditChartPanels(chart1);
            panelPage.IsPanelExist(chart1.DisplayName);

            //Post-Condition
            panelPage.DeletePanels("All");

        }

        /// <summary>
        /// Verify that user is able to successfully edit "Folder" field with valid path.
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC054()
        {
            Console.WriteLine("DA_PANEL_TC054 - Verify that user is able to successfully edit \"Folder\" field with valid path");

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //4. Create a new panel
            Chart chart = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.AddNewPanel(chart);
            mainPage.ClickLinkText(page.PageName);

            //5. Click Choose Panel button
            //6. Click on the newly created panel link
            mainPage.OpenPanelConfigurationFromChoosePanel(chart.DisplayName);

            //7. Edit valid folder path
            //8. Click Ok button
            AddNewPanelPage addPanelPage = new AddNewPanelPage(_webDriver);
            addPanelPage.SettingPanel(page.PageName, 400, "/Music Library/Actions");

            //VP: User is able to successfully edit "Folder" field with valid path
            mainPage.OpenEditPanelPopup();
            string actualDisplays = addPanelPage.TxtFolder.Value;
            Assert.AreEqual("/Music Library/Actions", actualDisplays, "Actual Folder Path is: " + actualDisplays);

            //Post-Condition
            addPanelPage.ClosePanelDialog("Cancel");
            mainPage.DeletePage(page.PageName);
            mainPage.OpenPanelsPage();
            panelPage.DeletePanels(chart.DisplayName);
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that all changes made to or with the values populated for corresponding parameters under "Categories" and "Series" field in Edit Panel are recorded correctly
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC062()
        {
            Console.WriteLine("DA_PANEL_TC062 - Verify that all changes made to or with the values populated for corresponding parameters under \"Categories\" and \"Series\" field in Edit Panel are recorded correctly");

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click Choose Panels button
            //4. Click Test Module Implementation By Priority link
            //5. Click Ok button on Panel Configuration dialog
            mainPage.OpenPanelConfigurationFromChoosePanel("Test Module Implementation By Priority").BtnOKConfigurationPanel.Click();

            //6. Click Edit Panel icon
            //7. Enter value into Caption field for Category
            //8. Enter value into Caption field for Serius
            //9. Click Ok button
            Chart chart1 = new Chart(null, null, null, 400, null, "Chart@","Stacked Bar", "Name", "Catagory Caption", "Location","Series Caption", null, null, null, false);
            mainPage.ClickEditPanelIcon("Test Module Implementation By Priority").EditChartPanel(chart1);

            //10. Click Edit Panel icon
            //VP. Caption's values are saved
            AddNewPanelPage editPanelPage = mainPage.ClickEditPanelIcon("Test Module Implementation By Priority");
            Assert.AreEqual("Catagory Caption", editPanelPage.TxtCategoryCaption.Value, string.Format("The current value is {0}", editPanelPage.TxtCategoryCaption.Value));
            Assert.AreEqual("Series Caption", editPanelPage.TxtSeriesCaption.Value, string.Format("The current value is {0}", editPanelPage.TxtSeriesCaption.Value));

        }

        /// <summary>
        /// Verify that for "Action Implementation By Status" panel instance, when user changes from "Pie" chart to any other chart type then change back the "Edit Panel" form should be as original
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC063()
        {
            Console.WriteLine("DA_PANEL_TC063 - Verify that for \"Action Implementation By Status\" panel instance, when user changes from \"Pie\" chart to any other chart type then change back the \"Edit Panel\" form should be as original");

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);


            //3. Click Choose Panels button
            //4. Click Action Implementation By Status link
            //5. Click Ok button on Panel Configuration dialog
            mainPage.OpenPanelConfigurationFromChoosePanel("Action Implementation By Status").BtnOKConfigurationPanel.Click();

            //6. Click Edit Panel icon
            //7. Click on Chart Type dropped down menu
            //8. Select Single Bar
            //9. Click on Chart Type dropped down menu
            //10. Select Pie
            //VP. Check original "Pie" - Edit Panel form is displayed
            //12. Close "Edit Panel" form
            AddNewPanelPage editPanelPage = mainPage.ClickEditPanelIcon("Action Implementation By Status");
            editPanelPage.FillPanelData(null, null, null, null, null, "Single Bar", null, null, null, null, false, null);
            editPanelPage.FillPanelData(null, null, null, null, null, "Pie", null, null, null, null, false, null);
            editPanelPage.BtnCancel.Click();

            //13. Click Edit Panel icon
            //14. Click on Chart Type dropped down menu
            //15. Select Stacked Bar
            //16. Click on Chart Type dropped down menu
            //17. Select Pie
            //VP. Check original "Pie" - Edit Panel form is displayed
            //18. Close "Edit Panel" form
            //19. Click Edit Panel icon
            mainPage.ClickEditPanelIcon("Action Implementation By Status");
            editPanelPage.FillPanelData(null, null, null, null, null, "Stacked Bar", null, null, null, null, false, null);
            editPanelPage.FillPanelData(null, null, null, null, null, "Pie", null, null, null, null, false, null);
            Assert.AreEqual(editPanelPage.CbbChartType.GetSelectedText(), "Pie", string.Format("Current selected value is {0}", editPanelPage.CbbChartType.GetSelectedText()));
            editPanelPage.BtnCancel.Click();

            //20. Click on Chart Type dropped down menu
            //21. Select Group Bar
            //22. Click on Chart Type dropped down menu
            //23. Select Pie
            //VP. Check original "Pie" - Edit Panel form is displayed
            //24. Close "Edit Panel" form
            mainPage.ClickEditPanelIcon("Action Implementation By Status");
            editPanelPage.FillPanelData(null, null, null, null, null, "Stacked Bar", null, null, null, null, false, null);
            editPanelPage.FillPanelData(null, null, null, null, null, "Pie", null, null, null, null, false, null);
            Assert.AreEqual(editPanelPage.CbbChartType.GetSelectedText(), "Pie", string.Format("Current selected value is {0}", editPanelPage.CbbChartType.GetSelectedText()));
            editPanelPage.BtnCancel.Click();

            //25. Click Edit Panel icon
            //26. Click on Chart Type dropped down menu
            //27. Select Line
            //28. Click on Chart Type dropped down menu
            //29. Select Pie
            //VP. Check original "Pie" - Edit Panel form is displayed
            mainPage.ClickEditPanelIcon("Action Implementation By Status");
            editPanelPage.FillPanelData(null, null, null, null, null, "Line", null, null, null, null, false, null);
            editPanelPage.FillPanelData(null, null, null, null, null, "Pie", null, null, null, null, false, null);
            Assert.AreEqual(editPanelPage.CbbChartType.GetSelectedText(), "Pie", string.Format("Current selected value is {0}", editPanelPage.CbbChartType.GetSelectedText()));

        }

        /// <summary>
        /// Verify that "Check All/Uncheck All" links are working correctly.
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_PANEL_TC064()
        {
            Console.WriteLine("DA_PANEL_TC064 - Verify that \"Check All/Uncheck All\" links are working correctly.");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);


            //1. Navigate to Dashboard login page
            //2. Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click 'Login' button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click 'Add Page' button
            //6. Enter Page Name
            //7. Click 'OK' button
            mainPage.AddPage(page);

            //8. Click 'Choose Panels' button below 'main_hung' button
            //9. Click 'Create new panel' button
            //10. Enter a name to Display Name
            //11. Click OK button
            //12. Click Cancel button
            mainPage.GoToPage(page.PageName);
            string panelName1 = CommonAction.GeneratePanelName();
            AddNewPanelPage addNewPanelPage = mainPage.OpenAddNewPanelPageFromButton().FillPanelData(null, null, panelName1, null, null, null, null, null, null, null, false, null);
            addNewPanelPage.BtnOk.Click();
            addNewPanelPage.BtnCancelConfigurationPanel.Click();

            //13. Click 'Create new panel' button
            //14. Enter a name to Display Name
            //15. Click OK button
            //16. Click Cancel button
            string panelName2 = CommonAction.GeneratePanelName();
            mainPage.OpenAddNewPanelPageFromButton().FillPanelData(null, null, panelName2, null, null, null, null, null, null, null, false, null);

            //17. Click 'Administer' link
            //18. Click 'Panels' link
            //19. Click 'Check All' link
            //VP. Check that 'hung_a' checkbox and 'hung_b' checkbox are checked
            //20. Click 'Uncheck All' link
            //VP. Check that 'hung_a' checkbox and 'hung_b' checkbox are unchecked
            string[] panelList = { panelName1, panelName2 };
            PanelsPage panelsPage = mainPage.OpenPanelsPage();
            panelsPage.LnkCheckAll.Click();
            Assert.IsTrue(panelsPage.IsCheckBoxExists(panelList), "All checkboxes are not checked");
            panelsPage.LnkUnCheckAll.Click();
            Assert.IsFalse(panelsPage.IsCheckBoxExists(panelList), "All checkboxes are not unchecked");

            //Post-Condition
            panelsPage.DeletePanels(panelName1);
            panelsPage.DeletePanels(panelName2);

        }

        /// <summary>
        /// Verify that user is unable to edit "Folder" field with invalid path
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC055()
        {
            Console.WriteLine("DA_PANEL_TC055 - Verify that user is unable to edit \"Folder\" field with invalid path");

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //4. Create a new panel
            Chart chart = new Chart(null, CommonAction.GeneratePanelName(), null, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.AddNewPanel(chart);
            mainPage.ClickLinkText(page.PageName);

            //5. Click Choose Panel button
            //6. Click on the newly created panel link
            mainPage.OpenPanelConfigurationFromChoosePanel(chart.DisplayName);

            //7. Edit valid folder path
            //8. Click Ok button
            AddNewPanelPage addPanelPage = new AddNewPanelPage(_webDriver);
            string actualMessage = addPanelPage.SettingPanelWithExpectedError(page.PageName, 400, "/Car Rental/Action");

            //VP: User is unable to edit "Folder" field with invalid path
            Assert.AreEqual("Panel folder is incorrect",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            addPanelPage.ConfirmDialog("OK");

            //Post-Condition
            addPanelPage.ClosePanelDialog("Cancel");
            mainPage.DeletePage(page.PageName);
            mainPage.OpenPanelsPage();
            panelPage.DeletePanels(chart.DisplayName);
        }

        /// <summary>
        /// Verify that user is unable to edit "Folder" field with empty value
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_PANEL_TC056()
        {
            Console.WriteLine("DA_PANEL_TC056 - Verify that user is unable to edit \"Folder\" field with empty value");

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Create a new page
            string pageName = CommonAction.GeneratePageName();
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            //4. Create a new panel
            Chart chart = new Chart(null, CommonAction.GeneratePanelName(), page.PageName, 400, null, "Chart@", null, null, null, "Name", null, null, null, null, false);
            AddNewPanelPage addPanelPage = new AddNewPanelPage(_webDriver);
            mainPage.OpenNewPanelPopUp(page.PageName);
            addPanelPage.AddChart(chart);
            mainPage.ClickLinkText(page.PageName);

            //5. Click Choose Panel button
            //6. Click on the newly created panel link
            mainPage.OpenPanelConfigurationFromChoosePanel(chart.DisplayName);

            //7. Edit valid folder path
            //8. Click Ok button
            string actualMessage = addPanelPage.SettingPanelWithExpectedError(page.PageName, 400, "");

            //VP: User is unable to edit "Folder" field with invalid path
            Assert.AreEqual("Panel folder is incorrect",
                actualMessage,
                string.Format("Failed! Actual message is: {0}", actualMessage));
            addPanelPage.ConfirmDialog("OK");

            //Post-Condition
            addPanelPage.ClosePanelDialog("Cancel");
            mainPage.DeletePage(page.PageName);
            PanelsPage panel = mainPage.OpenPanelsPage();
            panel.DeletePanels(chart.DisplayName);

        }
        /// <summary>
        /// DA_PANEL_TC057 - Verify that user is able to successfully edit \"Chart Type\"
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/06/2016</date>
        [TestMethod]
        public void DA_PANEL_TC057()
        {
            Console.WriteLine("DA_PANEL_TC057 - Verify that user is able to successfully edit \"Chart Type\"");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Create a new panel
            //7 Click Edit link
            //8 Change Chart Type for panel
            //9 Click Ok button            

            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);
            chart.ChartType = "Pie";
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.AddNewPanel(chart);
            chart.ChartType = "Single Bar";

            //VP Observe the current page: User is able to edit Chart Type successfully
            panelPage.EditChartPanels(chart);
            AddNewPanelPage newPanelPage = panelPage.OpenEditPanelPopup(chart.DisplayName);
            string chartType = newPanelPage.CbbChartType.GetSelectedText();
            Assert.AreEqual(chart.ChartType, chartType, "Cannot edit chart type for panel");

            newPanelPage.BtnCancel.Click();
            panelPage.DeletePanels(chart.DisplayName).Logout();

        }
        /// <summary>
        /// DA_PANEL_TC058 - Verify that \"Category\", \"Series\" and \"Caption\" field are enabled and disabled correctly corresponding to each type of the \"Chart Type\" in \"Edit Panel\" form
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/06/2016</date>
        [TestMethod]
        public void DA_PANEL_TC058()
        {
            Console.WriteLine("DA_PANEL_TC058 - Verify that \"Category\", \"Series\" and \"Caption\" field are enabled and disabled correctly corresponding to each type of the \"Chart Type\" in \"Edit Panel\" form");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Create a new panel
            //7 Click Edit link
            //8 Change Chart Type for panel: Pie
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);
            chart.ChartType = "Pie";
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            panelPage.AddNewPanel(chart);
            AddNewPanelPage newPanelPage = panelPage.OpenEditPanelPopup(chart.DisplayName);

            //VP Observe the current page -Category and Caption are disabled- Series is enabled
            Assert.AreEqual(true, newPanelPage.CbbSeries.Enabled, "Failed:  Series combobox are disabled");
            Assert.AreEqual(false, newPanelPage.CbbCategory.Enabled, "Failed: Category combobox are enable");
            Assert.AreEqual(false, newPanelPage.TxtCategoryCaption.Enabled, "Failed: Category are enable");
            Assert.AreEqual(false, newPanelPage.TxtSeriesCaption.Enabled, "Failed: Category are enable");

            //10 Change Chart Type for panel - Single Bar
            newPanelPage.CbbChartType.SelectByText("Single Bar");

            //VP Observe the current page
            //Category is disabled
            //Series and Caption are enabled
            //Assert.AreEqual(true,newPanelPage.ChbSeries.Enabled,)
            Assert.AreEqual(true, newPanelPage.CbbSeries.Enabled, "Failed:  Series combobox are disabled");
            Assert.AreEqual(false, newPanelPage.CbbCategory.Enabled, "Failed: Category combobox are enable");
            Assert.AreEqual(true, newPanelPage.TxtCategoryCaption.Enabled, "Failed: Category caption are disabled");
            Assert.AreEqual(true, newPanelPage.TxtSeriesCaption.Enabled, "Failed: Series caption are disabled");

            //12 Change Chart Type for panel
            newPanelPage.CbbChartType.SelectByText("Stacked Bar");

            //VP Observe the current page - All of them are enabled
            Assert.AreEqual(true, newPanelPage.CbbSeries.Enabled, "Failed: Series combobox are disabled");
            Assert.AreEqual(true, newPanelPage.CbbCategory.Enabled, "Failed: Category combobox are disabled");
            Assert.AreEqual(true, newPanelPage.TxtCategoryCaption.Enabled, "Failed: Category caption are disabled");
            Assert.AreEqual(true, newPanelPage.TxtSeriesCaption.Enabled, "Failed: Series caption are disabled");

            //14 Change Chart Type for panel
            newPanelPage.CbbChartType.SelectByText("Group Bar");

            //VP Observe the current page - All of them are enabled
            Assert.AreEqual(true, newPanelPage.CbbSeries.Enabled, "Failed:  Series combobox are disabled");
            Assert.AreEqual(true, newPanelPage.CbbCategory.Enabled, "Failed: Category combobox are disabled");
            Assert.AreEqual(true, newPanelPage.TxtCategoryCaption.Enabled, "Failed: Category caption are disabled");
            Assert.AreEqual(true, newPanelPage.TxtSeriesCaption.Enabled, "Failed: Series caption are disabled");

            //16 Change Chart Type for panel
            newPanelPage.CbbChartType.SelectByText("Line");

            //VP Observe the current page - All of them are enabled
            Assert.AreEqual(true, newPanelPage.CbbSeries.Enabled, "Failed:  Series combobox are disabled");
            Assert.AreEqual(true, newPanelPage.CbbCategory.Enabled, "Failed: Category combobox are disabled");
            Assert.AreEqual(true, newPanelPage.TxtCategoryCaption.Enabled, "Failed: Category caption are disabled");
            Assert.AreEqual(true, newPanelPage.TxtSeriesCaption.Enabled, "Failed: Series caption are disabled");

            newPanelPage.BtnCancel.Click();
            panelPage.DeletePanels(chart.DisplayName).Logout();
        }

        /// <summary>
        /// Verify that all settings within "Add New Panel" and "Edit Panel" form stay unchanged when user switches between "2D" and "3D" radio buttons in "Edit Panel" form
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/08/2016</date>
        [TestMethod]
        public void DA_PANEL_TC059()
        {
            Console.WriteLine("DA_PANEL_TC059 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"2D\" and \"3D\" radio buttons in \"Edit Panel\" form");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Switch between "2D" and "3D"
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);
            chart.ChartType = "Pie";
            chart.Style = "3D";
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            AddNewPanelPage addNewPanelPage = panelPage.OpenAddNewPanelPopupFromLink();
            addNewPanelPage.Rb3D.Click();

            //VP Observe the current page - All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //8 Create a new panel
            //9 Click Edit link
            //10 Switch between "2D" and "3D"
            addNewPanelPage.AddChart(chart);
            AddNewPanelPage newPanelPage = panelPage.OpenEditPanelPopup(chart.DisplayName);
            addNewPanelPage.Rb2D.Click();

            //VP Observe the current page - All settings is unchanged in the Edit Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            newPanelPage.BtnCancel.Click();
            panelPage.DeletePanels(chart.DisplayName).Logout();            
        }
        /// <summary>
        /// DA_PANEL_TC060 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"Legends\" radio buttons in \"Edit Panel\" form
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/08/2016</date>
        [TestMethod]
        public void DA_PANEL_TC060()
        {
            Console.WriteLine("DA_PANEL_TC060 - Verify that all settings within \"Add New Panel\" and \"Edit Panel\" form stay unchanged when user switches between \"Legends\" radio buttons in \"Edit Panel\" form");

            //1 Navigate to Dashboard login page
            //2 Login with valid account
            //3 Click Administer link
            //4 Click Panel link
            //5 Click Add New link
            //6 Click None radio button for Legends
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            Chart chart = new Chart(CommonAction.GeneratePanelName(), "Name", null);
            chart.ChartType = "Pie";
            PanelsPage panelPage = mainPage.OpenPanelsPage();
            AddNewPanelPage addNewPanelPage = panelPage.OpenAddNewPanelPopupFromLink();
            addNewPanelPage.RbNone.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //8 Click Top radio button for Legends
            addNewPanelPage.RbTop.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //10 Click Right radio button for Legends
            addNewPanelPage.RbRight.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //12 Click Bottom radio button for Legends
            addNewPanelPage.RbBottom.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //14 Click Left radio button for Legends
            addNewPanelPage.RbLeft.Click();
            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //16 Create a new panel Display Name: Logigear / Chart Type: Pie
            //17 Click Edit link
            //18 Click None radio button for Legends
            addNewPanelPage.AddChart(chart);
            AddNewPanelPage newPanelPage = panelPage.OpenEditPanelPopup(chart.DisplayName);
            addNewPanelPage.RbNone.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //20 Click Top radio button for Legends
            addNewPanelPage.RbTop.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //22 Click Right radio button for Legends
            addNewPanelPage.RbRight.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //24 Click Bottom radio button for Legends
            addNewPanelPage.RbBottom.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            //26 Click Left radio button for Legends
            addNewPanelPage.RbLeft.Click();

            //VP  Observe the current page --All settings is unchanged in the Add New Panel form
            Assert.AreEqual(true, addNewPanelPage.CbbDataProfile.Enabled, "Failed: Data Profile are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtDisplayName.Enabled, "Failed: Display Name are disabled");
            Assert.AreEqual(true, addNewPanelPage.TxtChartTitle.Enabled, "Failed: Chart Title are disabled");
            Assert.AreEqual(true, addNewPanelPage.CbbChartType.Enabled, "Failed: Chart Type are disabled");
            Assert.AreEqual(false, addNewPanelPage.CbbCategory.Enabled, "Failed: Category are enable");
            Assert.AreEqual(true, addNewPanelPage.CbbSeries.Enabled, "Failed: Series are disabled");
            Assert.AreEqual(false, addNewPanelPage.TxtCategoryCaption.Enabled, "Failed: Category Caption are enable");
            Assert.AreEqual(false, addNewPanelPage.TxtSeriesCaption.Enabled, "Failed: SeriesCaption are enable");
            Assert.AreEqual(true, addNewPanelPage.ChbShowTitle.Enabled, "Failed: Show Title are disabled");

            newPanelPage.BtnCancel.Click();
            panelPage.DeletePanels(chart.DisplayName).Logout();            
        }
           
    }
}
