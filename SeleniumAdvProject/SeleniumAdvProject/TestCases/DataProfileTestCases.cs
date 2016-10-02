using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumAdvProject.PageObjects;
using SeleniumAdvProject.Common;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using System.Text.RegularExpressions;
using SeleniumAdvProject.DataObjects;

namespace SeleniumAdvProject.TestCases
{
    [TestClass]
    public class DataProfileTestCases : BaseTestCase
    {
        /// <summary>
        /// Verify that all Pre-set Data Profiles are populated correctly
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC065()
        {
            Console.WriteLine("DA_DP_TC065 - Verify that all Pre-set Data Profiles are populated correctly");

            //1. Navigate to Dashboard login page
            //2. Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click Login
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click Administer->Data Profiles
            //VP. Check Pre-set Data Profile are populated correctly in profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();
            Assert.IsTrue(DPPage.IsPresetDataProfilePopulated(), "Pre-set Data Profile aren't populated correctly");

        }

        /// <summary>
        /// Verify that all Pre-set Data Profiles are populated correctly
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC066()
        {
            Console.WriteLine("DA_DP_TC066 - Verify that all Pre-set Data Profiles are populated correctly");

            //1. Navigate to Dashboard login page
            //2. Select a specific repository 
            //3. Enter valid Username and Password
            //4. Click Login
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //5. Click Administer->Data Profiles
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //VP. Check there is no 'Delele' or 'Edit' link appears in Action section of Pre-set Data Profiles
            //6. Click on Pre-set Data Profile name
            //VP. Check there is no link on Pre-set Data Profile name
            //VP. Check there is no checkbox appears in the left of Pre-set Data Profiles
            foreach (string presetDP in Constants.presetDataProfile)
            {
                Assert.IsFalse(DPPage.IsLinkExists(presetDP, "Edit"), string.Format("Edit link is exists for preset data profile \"{0}\"", presetDP));
                Assert.IsFalse(DPPage.IsLinkExists(presetDP, "Delete"), string.Format("Delete link is exists for preset data profile \"{0}\"", presetDP));
                Assert.IsFalse(DPPage.IsLinkExists(presetDP, presetDP), string.Format("{0} link is exists for preset data profile \"{0}\"", presetDP));
                Assert.IsFalse(DPPage.IsCheckBoxExists(presetDP), string.Format("CheckBox is exists for preset data profile \"{0}\"", presetDP));
            }
        }

        /// <summary>
        /// Verify that all fields are displayed correctly
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC077()
        {
            Console.WriteLine("DA_DP_TC077 - Verify that all fields are displayed correctly");

            //1. Navigate to Dashboard login page
            //2. Log in specific repository with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click on "Administer" llink
            //4. Click on "Data Profiles" link
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //5. Click "Add New" link
            //6. Enter Name field
            //7. Click Item Type listbox
            //8. Select specific Item Type
            //9. Click on "Next" button
            //VP. Check all fields are displayed correctly 
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), null, null);
            string[] displayField = { "Name", "Description", "Source", "Assigned user", "Status", "Last updated by", "Created by", "Check out by", "Location", "Recent result", "Version", "Priority", "Last update date", "Creation date", "Notes", "URL" };
            bool actualresult = false;
            foreach (string field in displayField)
            {
                actualresult = displayFieldsPage.IsFieldDisplayed(field);
                if (!actualresult)
                    break;
            }
            Assert.IsTrue(actualresult, "Display Fields displays uncorrectly");
        }

        /// <summary>
        /// Verify that all fields are pre-fixed with check boxes
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC078()
        {
            Console.WriteLine("DA_DP_TC078 - Verify that all fields are pre-fixed with check boxes");

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on Data Profile "A"
            //4. Click on Next button
            //5. Check each field listed is prefixed with checkbox
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), null, null);
            string[] displayField = { "Name", "Description", "Source", "Assigned user", "Status", "Last updated by", "Created by", "Check out by", "Location", "Recent result", "Version", "Priority", "Last update date", "Creation date", "Notes", "URL" };
            bool actualresult = false;
            foreach (string field in displayField)
            {
                actualresult = displayFieldsPage.IsCheckBoxFieldDisplayed(field);
                if (!actualresult)
                    break;
            }
            Assert.IsTrue(actualresult, "Display Fields displays uncorrectly");

        }

        /// <summary>
        /// Verify that Check All / Uncheck All Links are working correctly
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC079()
        {
            Console.WriteLine("DA_DP_TC079 - Verify that Check All / Uncheck All Links are working correctly");

            //1. Navigate to Dashboard login page
            //2. Log in with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Click on "Administer" llink
            //4. Click on "Data Profiles" link
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //5. Click "Add New" link
            //6. Enter Name field
            //7. Click on "Next" button
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), null, null);

            //8. Click on "Check All" link
            //VP.Verify that all checkbox is checked
            displayFieldsPage.ClickCheckAll();
            string[] displayField = { "Name", "Description", "Source", "Assigned user", "Status", "Last updated by", "Created by", "Check out by", "Location", "Recent result", "Version", "Priority", "Last update date", "Creation date", "Notes", "URL" };
            bool actualresult = false;
            foreach (string field in displayField)
            {
                actualresult = displayFieldsPage.IsCheckBoxChecked(field);
                if (!actualresult)
                    break;
            }
            Assert.IsTrue(actualresult, "All checkboxes aren't checked");

            //9. Click on "Uncheck All" link
            //VP. Verify that all checkbox is unchecked
            displayFieldsPage.ClickUnCheckAll();
            foreach (string field in displayField)
            {
                actualresult = displayFieldsPage.IsCheckBoxChecked(field);
                if (actualresult)
                    break;
            }
            Assert.IsFalse(actualresult, "All checkboxes aren't unchecked");

        }

        /// <summary>
        /// Verify that all fields are listed in the "Field" dropped down menu
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_DP_TC080()
        {
            Console.WriteLine("DA_DP_TC080 - Verify that all fields are listed in the \"Field\" dropped down menu");

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Input to "Name *" field
            //4. Click "Item Type" dropped down menu and choose Test Modules
            //5. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test modules", null);
            SortFieldsPage SortFieldsPage = displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType1 = { "--- Select field ---", "Name", "Location", "Description", "Recent result", "Source", "Version", "Assigned user", "Priority", "Status", "Last update date", "Last updated by", "Creation date", "Created by", "Notes", "Check out by", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType1), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //6. Navigate to General Settings page
            //7. Click "Item Type" dropped down menu and choose Test Cases
            //8. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test cases", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType2 = { "--- Select field ---", "ID", "Location", "Title", "Recent result", "Notes", "Source", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType2), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //9. Navigate to General Settings page
            //10. Click "Item Type" dropped down menu and choose Test Objectives
            //11. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test objectives", null);
            displayFieldsPage.GoToSortFieldsPage();
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType2), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //12. Navigate to General Settings page
            //13. Click "Item Type" dropped down menu and choose Data Sets
            //14. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "data sets", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType3 = { "--- Select field ---", "Name", "Location", "Description", "Version", "Assigned user", "Status", "Last update date", "Last updated by", "Creation date", "Created by", "Notes", "Check out by", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType3), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //15. Navigate to General Settings page
            //16. Click "Item Type" dropped down menu and choose Actions
            //17. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "actions", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType4 = { "--- Select field ---", "Name", "Location", "Description", "Version", "Assigned user", "Status", "Last update date", "Last updated by", "Creation date", "Created by", "Notes", "Check out by", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType4), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //18. Navigate to General Settings page
            //19. Click "Item Type" dropped down menu and choose Interface Entities
            //20. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "interface entities", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType5 = { "--- Select field ---", "Name", "Location", "Description", "Version", "Assigned user", "Status", "Last update date", "Last updated by", "Creation date", "Created by", "Notes", "Check out by", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType5), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //22. Navigate to General Settings page
            //23. Click "Item Type" dropped down menu and choose Test Results
            //24. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test results", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType6 = { "--- Select field ---", "Name", "Location", "Reported by", "Date of run", "Start time", "End time", "Duration", "Comment", "Variation", "Result", "Passed", "Failed", "Warnings", "Errors", "Automated/Manual", "Run Machine", "Notes", "URL", "Baseline", "OS", "Device OS", "Device Name", "Build Number", "AUTVersion" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType6), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

            //25. Navigate to General Settings page
            //26. Click "Item Type" dropped down menu and choose Test Case Results
            //27. Navigate to Sort Fields page
            //VP. Check all fields of selected "Item Type" item are listed under the "Field" dropped down menu 
            SortFieldsPage.GoToGeneralSettingsPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test case results", null);
            displayFieldsPage.GoToSortFieldsPage();
            string[] expectedItemType7 = { "--- Select field ---", "Name", "Location", "Description", "Version", "Assigned user", "Status", "Last update date", "Last updated by", "Creation date", "Created by", "Notes", "Check out by", "URL" };
            Assert.IsTrue(SortFieldsPage.IsItemTypeListed(expectedItemType7), "\"Item Type\" item aren't listed under the \"Field\" dropped down menu");

        }

        /// <summary>
        /// Verify that Data Profiles are listed alphabetically
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC067()
        {
            Console.WriteLine("DA_DP_TC067 - Verify that Data Profiles are listed alphabetically");

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Click Administer->Data Profiles
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //VP: Check Data Profiles are listed alphabetically
            bool actualResult = DPPage.IsDataProfileContentSorted("panel_tag1", "ASC");
            Assert.AreEqual(true, actualResult, "Data Profiles are not listed alphabetically");

            //Post-Condition
            mainPage.Logout();

        }

        /// <summary>
        /// Verify that Check Boxes are only present for non-preset Data Profiles.
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC068()
        {
            Console.WriteLine("DA_DP_TC068 - Verify that Check Boxes are only present for non-preset Data Profiles");

            //Set Variable
            string profileName = CommonAction.GenerateDataProfileName();

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Click Administer->Data Profiles
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Create a new Data Profile
            DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(profileName, null, null, "Finish");

            //VP: Check Check Boxes are only present for non-preset Data Profiles
            bool isProfileCheckBox = DPPage.IsCheckBoxExists(profileName);
            Assert.AreEqual(true, isProfileCheckBox, "This Data Profile does not have checkbox");

            //Post-Condition
            DPPage.DeleteAllDataProfiles();
        }

        /// <summary>
        /// Verify that user is unable to proceed to next step or finish creating data profile if  "Name *" field is left empty
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC069()
        {
            Console.WriteLine("DA_DP_TC069 - Verify that user is unable to proceed to next step or finish creating data profile if  \"Name *\" field is left empty");

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on "Add New"
            //4. Click on "Next Button"
            //VP: dialog message "Please input profile name" appears
            string actualMessage = DPPage.GoToGeneralSettingPage().SetGeneralSettingsWithExpectedError(" ");
            Assert.AreEqual("Please input profile name.", actualMessage,
                           string.Format("Failed! Actual message is: {0}", actualMessage));
            DPPage.ConfirmDialog("OK");

            //5. Click on "Finish Button"
            //VP: dialog message "Please input profile name" appears
            GeneralSettingsPage generalPage = new GeneralSettingsPage(_webDriver);
            generalPage.BtnFinish.Click();
            string actualMessageFinish = generalPage.GetDialogText();
            Assert.AreEqual("Please input profile name.", actualMessageFinish,
                           string.Format("Failed! Actual message is: {0}", actualMessageFinish));
            DPPage.ConfirmDialog("OK");

            //Post Condition
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that special characters ' /:*?<>|"#[ ]{}=%; 'is not allowed for input to "Name *" field
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC070()
        {
            Console.WriteLine("DA_DP_TC070 - Verify that special characters ' /:*?<>|\"#[ ]{}=%; 'is not allowed for input to \"Name *\" field");

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on "Add New"
            //4. Input special character
            //5. Click on "Next Button"
            string actualMessage = DPPage.GoToGeneralSettingPage().SetGeneralSettingsWithExpectedError("a123:\"/{}!@$\u005c", "Next");

            //VP: Check dialog message indicates invalid characters: /:*?<>|"#[ ]{}=%; is not allowed as input for name field appears
            Assert.AreEqual("Invalid name. The name cannot contain high ASCII characters or any of the following characters: /:*?<>|\"#[]{}=%;"
                , actualMessage, string.Format("Failed! Actual message is: {0}", actualMessage));
            DPPage.ConfirmDialog("OK");

            //Post Condition
            mainPage.Logout();
        }
        /// <summary>
        /// Verify that Data Profile names are not case sensitive
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC071()
        {
            Console.WriteLine("DA_DP_TC071 - Verify that Data Profile names are not case sensitive");

            //Set Variables
            string nameProfile = CommonAction.GenerateDataProfileName();
            string uppercaseName = nameProfile.ToUpper();

            //1. Log in Dashboard
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //Pre-Condition: Create a data profile
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();
            DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(nameProfile, null, null, "Finish");

            //2. Navigate to Data Profiles page
            //3. Click on "Add New"
            //4. Input charater uppercase name into "Name *" field
            //5. Click "Next" button 
            string actualMessage = DPPage.GoToGeneralSettingPage().SetGeneralSettingsWithExpectedError(uppercaseName, "Next");

            //VP: Check dialog message "Data Profile name already exists"
            Assert.AreEqual("Data profile name already exists."
                , actualMessage, string.Format("Failed! Actual message is: {0}", actualMessage));
            DPPage.ConfirmDialog("OK");

            //Post-condition
            mainPage.GoToDataProfilePage();
            DPPage.DeleteAllDataProfiles();
        }
        /// <summary>
        /// DA_DP_TC072 - Verify that all data profile types are listed under \"Item Type\" dropped down menu
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/08/2016</date>
        [TestMethod]
        public void DA_DP_TC072()
        {
            Console.WriteLine("DA_DP_TC072 - Verify that all data profile types are listed under \"Item Type\" dropped down menu");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click Login
            //5 Click Administer->Data Profiles
            //6 Click 'Add New' link
            string[] itemsList = new string[] { "Test Modules", "Test Cases", "Test Objectives", "Data Sets", "Actions", "Interface Entities", "Test Results", "Test Case Results" };
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            DataProfilePage dataProfilePage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password).GoToDataProfilePage();
            GeneralSettingsPage genralSettingPage = dataProfilePage.GoToGeneralSettingPage();

            //VP All data profile types are listed under "Item Type" dropped down menu
            //+ Test Modules
            //+ Test Cases
            //+ Test Objectives
            //+ Data Sets
            //+ Actions
            //+ Interface Entities
            //+ Test Results
            //+ Test Case Results
            Assert.AreEqual(true, genralSettingPage.isComboboxContainsItems(genralSettingPage.CbbItemType, itemsList));

            genralSettingPage.Logout();
        }

        /// <summary>
        /// DA_DP_TC073 - Verify that all data profile types are listed in priority order under \"Item Type\" dropped down menu
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/08/2016</date>
        [TestMethod]
        public void DA_DP_TC073()
        {
            Console.WriteLine("DA_DP_TC073 - Verify that all data profile types are listed in priority order under \"Item Type\" dropped down menu");

            //1 Log in Dashboard
            //2 Navigate to Data Profiles page
            //3 Click on "Add New"
            //4 Click on "Item Type" dropped down menu          

            string[] itemsList = new string[] { "Test Modules", "Test Cases", "Test Objectives", "Data Sets", "Actions", "Interface Entities", "Test Results", "Test Case Results" };
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            GeneralSettingsPage genralSettingPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .GoToDataProfilePage().GoToGeneralSettingPage();

            //VP "Item Type" items are listed in priority order: Test Modules>Test Cases> Test Objectives> Data Sets> Actions> Interface Entities> Test Results> Test Cases results
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedByPriorityOrder(genralSettingPage.CbbItemType, itemsList), "Failed! \"Item Type\" items are not listed in priority order");

            genralSettingPage.Logout();
        }
        /// <summary>
        /// DA_DP_TC074 - Verify that appropriate \"Related Data\" items are listed correctly corresponding to the \"Item Type\" items.
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/09/2016</date>
        [TestMethod]
        public void DA_DP_TC074()
        {
            Console.WriteLine("DA_DP_TC074 - Verify that appropriate \"Related Data\" items are listed correctly corresponding to the \"Item Type\" items.");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click Login
            //5 Click Administer->Data Profiles
            //6 Click Add new link
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            GeneralSettingsPage genralSettingPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .GoToDataProfilePage().GoToGeneralSettingPage();

            //7 Select 'Test Modules' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Test Modules");

            //VP Check 'Related Data' items listed correctly {Related Test Results,Related Test Cases}
            string[] expectedList = new string[] { "Related Test Results", "Related Test Cases" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //9 Select 'Test Cases' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Test Cases");

            //VP Check 'Related Data' items listed correctly {Related Run Results,Related Objectives}
            expectedList = new string[] { "Related Test Results", "Related Objectives" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //11 Select 'Test Objectives' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Test Objectives");

            //VP Check 'Related Data' items listed correctly {Related Run Results,Related Test Cases}
            expectedList = new string[] { "Related Test Results", "Related Objectives" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //13 Select 'Data Sets' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Data Sets");

            //VP Check 'Related Data' items listed correctly {None}
            expectedList = new string[] { "None" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //15 Select 'Actions' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Actions");

            //VP Check 'Related Data' items listed correctly {Action Arguments}
            expectedList = new string[] { "Action Arguments" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //17 Select 'Interface Entities' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Interface Entities");

            //VP Check 'Related Data' items listed correctly {Interface Elements}
            expectedList = new string[] { "Interface Elements" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //19 Select 'Test Results' in 'Item Type' drop down list 
            genralSettingPage.CbbItemType.SelectByText("Test Results");

            //VP Check 'Related Data' items listed correctly {Related Test Module, Related Test Cases}
            expectedList = new string[] { "Related Test Results", "Related Test Cases" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            //21 Select 'Test Case Results' in 'Item Type' drop down list
            genralSettingPage.CbbItemType.SelectByText("Test Case Results");

            //VP Check 'Related Data' items listed correctly
            expectedList = new string[] { "None" };
            Assert.AreEqual(true, genralSettingPage.CheckItemsInComboboxListedCorrectly(genralSettingPage.CbbRelatedData, expectedList), "Failed!Check 'Related Data' items did not list correctly");

            genralSettingPage.Logout();
        }
        /// <summary>
        /// DA_DP_TC075 - Verify that default settings are applied correctly for newly created data profiles if user only set up \"General Settings\" page and finishes.
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/09/2016</date>
        [TestMethod]
        public void DA_DP_TC075()
        {
            Console.WriteLine("DA_DP_TC075 - Verify that default settings are applied correctly for newly created data profiles if user only set up \"General Settings\" page and finishes.");

            //1 Navigate to Data Profiles page
            //2 Login with valid account
            //3 Click on "Add New"
            //4 Input to "Name *" field
            //5 Click "Item Type" and choose an item
            //6 Click "Finish" button
            string profileName = CommonAction.GenerateDataProfileName();
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();
            DPPage = (DataProfilePage)DPPage.GoToGeneralSettingPage()
                .SetGeneralSettingsValue(profileName, "test modules", "None", "Finish");

            //7 Click on the newly created data profile
            DPPage.ClickLinkText(profileName);
            //VP Check the setting of General Settings Page
            Assert.AreEqual("General Settings", DPPage.LblProfileHeader.Text, "Failed! Page {0}");


            //9 Click Next Button
            //10 Check the setting of Display Fields Page
            //11 Click Next Button
            //12 Check the setting of Sort Fields Page
            //13 Click Next Button
            //14 Check the setting of Filter Fields Page
            //15 Click Next Button
            //16 Check the setting of Statistic Page


        }
        /// <summary>
        /// DA_DP_TC076 - Verify that for newly created data profile, user is able to navigate through other setting pages on the left navigation panel.
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/10/2016</date>
        [TestMethod]
        public void DA_DP_TC076()
        {
            Console.WriteLine("DA_DP_TC076 - Verify that for newly created data profile, user is able to navigate through other setting pages on the left navigation panel.");

            //1 Navigate to Dashboard login page
            //2 Select a specific repository 
            //3 Enter valid Username and Password
            //4 Click Login
            //5 Click Administer->Data Profiles
            //6 Click Add new link
            //7 Create new data profile
            //8 Back to Data Profiles page
            //9 Click on the data profile 'thinh-test'
            string profileName = CommonAction.GenerateDataProfileName();
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();
            DPPage = (DataProfilePage)DPPage.GoToGeneralSettingPage()
                .SetGeneralSettingsValue(profileName, "test modules", "None", "Finish");
            DPPage.ClickLinkText(profileName);

            //10 Click on 'Display Fields' in the left navigation panel
            DPPage.LblDisplayField.Click();

            //11 Check Display Fields page appears
            Assert.AreEqual("General Settings", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of General Setting", DPPage.LblProfileHeader.Text);

            //12 Click on 'Sort Fields' in the left navigation panel
            DPPage.LblSortField.Click();

            //13 Check Sort Fields page appears
            Assert.AreEqual("Sort Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Sort Fields", DPPage.LblProfileHeader.Text);

            //14 Click on 'Filter Fields' in the left navigation panel
            DPPage.LblFilterFields.Click();

            //15 Check Filter Fields page appears
            Assert.AreEqual("Filter Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of 'Filter Fields", DPPage.LblProfileHeader.Text);

            //16 Click on 'Statistic Fields' in the left navigation panel
            DPPage.LblStatisticFields.Click();

            //17 Check Statistic Fields page appears
            Assert.AreEqual("Statistic Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Statistic Fields", DPPage.LblProfileHeader.Text);

            //18 Click on 'Display Sub-Fields' in the left navigation panel
            DPPage.LblDisplaySubFiled.Click();

            //19 Check Display Sub-Fields page appears
            Assert.AreEqual("Display Sub-Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Display Sub-Fields", DPPage.LblProfileHeader.Text);

            //20 Click on 'Sort Sub-Fields' in the left navigation panel
            DPPage.LblSortSubField.Click();

            //21 Check Sort Sub-Fields page appears
            Assert.AreEqual("Sort Sub-Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Sort Sub-Fields", DPPage.LblProfileHeader.Text);
            //22 Click on 'Filter Sub-Fields' in the left navigation panel
            DPPage.LblFilterSubField.Click();

            //23 Check Filter Sub-Fields page appears
            Assert.AreEqual("Filter Sub-Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Filter Sub-Fields", DPPage.LblProfileHeader.Text);

            //24 Click on 'Statistic Sub-Fields' in the left navigation panel
            DPPage.LblStatisticFields.Click();

            //25 Check Statistic Sub-Fields page appears
            Assert.AreEqual("Statistic Sub-Fields", DPPage.LblProfileHeader.Text,
                "Failed! Page {0} display instead of Statistic Sub-Fields", DPPage.LblProfileHeader.Text);

            DPPage.GoToDataProfilePage();
            DPPage.DeleteDataProfiles(profileName);

        }

        /// <summary>
        /// Verify that user is unable to add any field more than once.
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC084()
        {
            Console.WriteLine("DA_DP_TC084 - Verify that user is unable to add any field more than once.");

            //1. Log in Dashboard 
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on "Add New"
            //4. Input to "Name *" field
            //5. Click on Next button
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test modules", null);

            //6. Navigate to Sort Fields page
            SortFieldsPage sortFieldsPage = displayFieldsPage.GoToSortFieldsPage();

            //7. Click on "Field" dropped down menu
            //8. Select "Name" item
            //9. Click on "Add Level" button
            sortFieldsPage.AddLevel("Name");

            //VP: Check "Name item is added to the sorting criteria list
            bool actualField = sortFieldsPage.IsFieldLevelExist("Name");
            Assert.AreEqual(true, actualField, "Field Level: " + actualField + " does not exist");

            //10. Click on "Field" dropped down menu
            //11. Select "Name" item
            //12. Click on "Add Level" button
            string warningMessage = sortFieldsPage.AddLevelWithExpectedError("Name");
            Assert.AreEqual("Field 'Name' already selected"
                , warningMessage, string.Format("Failed! Actual message is: {0}", warningMessage));
            sortFieldsPage.ConfirmDialog("OK");

            //Post-Condition
            sortFieldsPage.RemoveFieldLevel("Name");
        }

        /// <summary>
        /// Verify that all date types are listed correctly under date type dropped down menu
        /// </summary>
        /// Author: Tu Nguyen
        [TestMethod]
        public void DA_DP_TC087()
        {
            Console.WriteLine("DA_DP_TC087 - Verify that all date types are listed correctly under date type dropped down menu");

            //Set Variables
            string[] itemsList = new string[] { "day", "week", "month", "year" };

            //1. Log in Dashboard 
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on "Add New"
            //4. Input to "Name *" field
            //5. Click on Next button
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test modules", null);

            //6. Navigate to Sort Fields page
            SortFieldsPage sortFieldsPage = displayFieldsPage.GoToSortFieldsPage();

            //7. Click on "Field" dropped down menu
            //8. Select "Last update date" item
            //9. Click on "Add Level" button
            sortFieldsPage.AddLevel("Last update date");
            bool actualContain = sortFieldsPage.isComboboxContainsItems(sortFieldsPage.CbbDateField, itemsList);

            //VP: all date types from date type  dropped down menu are listed correctly
            Assert.AreEqual(true, actualContain, "Item does not exist");

            //Post-Condition
            sortFieldsPage.RemoveFieldLevel("Last update date");
        }

        /// <summary>
        /// Verify that user is able to add levels of fields
        /// <author>Tu Nguyen</author>
        [TestMethod]
        public void DA_DP_TC082()
        {
            Console.WriteLine(" DA_DP_TC082 - Verify that user is able to add levels of fields ");

            //1. Log in Dashboard 
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //2. Navigate to Data Profiles page
            DataProfilePage DPPage = mainPage.GoToDataProfilePage();

            //3. Click on "Add New"
            //4. Input to "Name *" field
            //5. Click on Next button
            DisplayFieldsPage displayFieldsPage = (DisplayFieldsPage)DPPage.GoToGeneralSettingPage().SetGeneralSettingsValue(CommonAction.GenerateDataProfileName(), "test modules", null);

            //6. Navigate to Sort Fields page
            SortFieldsPage sortFieldsPage = displayFieldsPage.GoToSortFieldsPage();

            //7. Click on "Field" dropped down menu
            //8. Select "Name" item
            //9. Click on "Add Level" button
            sortFieldsPage.AddLevel("Name");

            //VP: Check "Name item is added to the sorting criteria list
            bool actualField = sortFieldsPage.IsFieldLevelExist("Name");
            Assert.AreEqual(true, actualField, "Field Level: " + actualField + " does not exist");

            //10. Click on "Field" dropped down menu
            //11. Select another item Location
            //12. Click on "Add Level" button
            sortFieldsPage.AddLevel("Location");
            actualField = sortFieldsPage.IsFieldLevelExist("Location");
            Assert.AreEqual(true, actualField, "Field Level: " + actualField + " does not exist");

            //Post-Condition
            sortFieldsPage.RemoveFieldLevel("Name").RemoveFieldLevel("Location");
        }
    }
}

