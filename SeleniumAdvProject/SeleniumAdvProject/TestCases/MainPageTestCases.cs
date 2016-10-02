using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumAdvProject.PageObjects;
using SeleniumAdvProject.Common;
using SeleniumAdvProject.DataObjects;
using System.Threading;
using OpenQA.Selenium;

namespace SeleniumAdvProject.TestCases
{
    [TestClass]
    public class MainPageTestCases : BaseTestCase
    {
        /// <summary>
        /// Verify that user is unable open more than 1 "New Page" dialog
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC011()
        {
            Console.WriteLine("DA_MP_TC011 - Verify that user is unable open more than 1 \"New Page\" dialog");

            //1 Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account
            //3. Go to Global Setting -> Add page  
            //4. Try to go to Global Setting -> Add page again
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .OpenSetting();

            //VP. User cannot go to Global Setting -> Add page while "New Page" dialog appears
            Assert.IsFalse(mainPage.IsSettingExist(), "User cannot go to Global Setting while \"New Page\" dialog appears");

        }

        /// <summary>
        /// Verify that user is able to add additional pages besides "Overview" page successfully
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC012()
        {
            Console.WriteLine("DA_MP_TC012 - Verify that user is able to add additional pages besides \"Overview\" page successfully");

            //Set variables
            string pageName = CommonAction.GenrateRandomString(Constants.lenghtRandomString);
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account
            //3. Go to Global Setting -> Add page
            //4. Enter Page Name field
            //5. Click OK button
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(page);

            //VP. Check "Test" page is displayed besides "Overview" page
            Assert.IsTrue(mainPage.IsPageDisplayAfter("Overview", page.PageName), "\"Test\" page is not displayed besides \"Overview\" page");

            //Post condition
            mainPage.DeletePage(page.PageName).Logout();
        }

        /// <summary>
        /// Verify that the newly added main parent page is positioned at the location specified 
        /// as set with "Displayed After" field of "New Page" form on the main page bar"Parent Page" dropped down menu
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC013()
        {
            Console.WriteLine("DA_MP_TC013 - Verify that the newly added main parent page is positioned at the location specified as set with \"Displayed After\" field of \"New Page\" form on the main page bar \"Parent Page\" dropped down menu");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, page.PageName, false);

            //1. Navigate to Dashboard login page
            //2. Log in specific repository with valid account
            //3. Go to Global Setting -> Add page            
            //4. Enter Page Name field
            //5. Click OK button
            //6. Go to Global Setting -> Add page
            //7. Enter Page Name field
            //8. Click on  Displayed After dropdown list
            //9. Select specific page
            //10. Click OK button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(page)
                .AddPage(page1);

            //VP. Check "Another Test" page is positioned besides the "Test" page                       
            Assert.IsTrue(mainPage.IsPageDisplayAfter(page.PageName, page1.PageName),
                "The {0} page is not beside the {1} page", page1.PageName, page.PageName);

            //Post-Condition
            mainPage.DeletePage(page.PageName)
                .DeletePage(page1.PageName)
                .Logout();
        }

        /// <summary>
        /// Verify that "Public" pages can be visible and accessed by all users of working repository
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC014()
        {
            Console.WriteLine("DA_MP_TC014 - Verify that \"Public\" pages can be visible and accessed by all users of working repository");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", true);

            //1. Navigate to Dashboard login page
            //2. Log in specific repository with valid account
            //3. Go to Global Setting -> Add page
            //4. Enter Page Name field
            //5. Check Public checkbox
            //6. Click OK button 
            //7. Click on Log out link   
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(page);
            loginPage = mainPage.Logout();

            //8. Log in with another valid account  
            //VP. Check newly added page is visibled
            mainPage = loginPage.Login(Constants.Repository, Constants.UserName1, Constants.SpecialPassword);
            Assert.IsTrue(!mainPage.IsPageExist(page.PageName), string.Format("{0} is not visibled", page.PageName));

            //Post-Condition
            mainPage.DeletePage(page.PageName).Logout();

        }

        /// <summary>
        /// Verify that non "Public" pages can only be accessed and visible to their creators with condition that all parent pages above it are "Public"
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC015()
        {
            Console.WriteLine("DA_MP_TC015 - Verify that non \"Public\" pages can only be accessed and visible to their creators with condition that all parent pages above it are \"Public\"");

            //Set variables
            Page parentPage = new Page("Test", "Select parent", 2, "Overview", false);
            Page childPage = new Page("TestChild", "Test", 2, "Select page", false);

            //1 Navigate to Dashboard login page
            //2 Log in specific repository with valid account
            //3. Go to Global Setting -> Add page
            //4 Enter Page Name field (Test)
            //5 Check Public checkbox
            //6 Click OK button
            //7 Go to Global Setting -> Add page
            //8 Enter Page Name field (Test Chilld)
            //9 Click on  Select Parent dropdown list
            //10 Select specific page (Test)
            //11 Click OK button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(parentPage)
                .AddPage(childPage);

            //12 Click on Log out link
            mainPage.Logout();

            //13 Log in with another valid account
            mainPage = loginPage.Login(Constants.Repository, "test", "admin");
            //mainPage = loginPage.Login(Constants.Repository, Constants.UserName1, Constants.SpecialPassword);

            //VP Check children is invisibled
            Assert.IsFalse(mainPage.IsPageExist(parentPage.PageName + "/" + childPage.PageName), "Failled by bug! child page is still displayed");

            //Post-Condition
            mainPage.DeletePage(parentPage.PageName + "/" + childPage.PageName)
                .DeletePage(parentPage.PageName)
                .Logout();
        }

        /// <summary>
        /// Verify that user is able to edit the "Public" setting of any page successfully
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC016()
        {
            Console.WriteLine("DA_MP_TC016 - Verify that user is able to edit the \"Public\" setting of any page successfully");

            //Set variables
            Page page1 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page page2 = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", true);

            //1 Navigate to Dashboard login page
            //2 Log in specific repository with valid account
            //3. Go to Global Setting -> Add page
            //4 Enter Page Name
            //5 Click OK button           
            //6 Go to Global Setting -> Add page
            //7 Enter Page Name
            //8 Check Public checkbox
            //9 Click OK button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password)
                .AddPage(page1)
                .AddPage(page2);

            //10 Click on "Test" page
            //11 Click on "Edit" link
            //12 Check "Edit Page" pop up window is displayed
            //13 Check Public checkbox
            //14 Click OK button
            //15 Click on "Another Test" page
            //16 Click on "Edit" link
            //17 Check "Edit Page" pop up window is displayed
            //18 Uncheck Public checkbox
            //19 Click OK button
            //20 Click Log out link
            //21 Log in with another valid account
            page1.IsPublic = true;
            page2.IsPublic = false;
            mainPage.OpenEditPage(page1.PageName);
            mainPage.EditPage(page1);
            mainPage.OpenEditPage(page2.PageName);
            mainPage.EditPage(page2)
                .Logout()
                .Login(Constants.Repository, Constants.UserName1, Constants.SpecialPassword);

            //VP Check "Test" Page is visible and can be accessed            
            Assert.AreEqual(true, mainPage.IsPageExist(page1.PageName));

            //VP Check "Another Test" page is invisible
            Assert.AreEqual(true, mainPage.IsPageExist(page2.PageName), string.Format("{0} page is visible", page2.PageName));

            //Post-Condition
            mainPage.DeletePage(page1.PageName)
                .Logout()
                .Login(Constants.Repository, Constants.UserName, Constants.Password)
                .DeletePage(page2.PageName)
                .Logout();
        }

        /// <summary>
        /// Verify that user can remove any main parent page except "Overview" page successfully and the order of pages stays persistent as long as there is not children page under it
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2016</date>
        [TestMethod]
        public void DA_MP_TC017()
        {
            Console.WriteLine("DA_MP_TC017 - Verify that user can remove any main parent page except \"Overview\" page successfully and the order of pages stays persistent as long as there is not children page under it");

            //Set variables
            Page parentPage = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page childPage = new Page(CommonAction.GeneratePageName(), parentPage.PageName, 3, "Select page", false);

            //1 Navigate to Dashboard login page
            //2 Log in specific repository with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Add a new parent page
            //4. Add a children page of newly added page
            mainPage.AddPage(parentPage);
            mainPage.AddPage(childPage);

            //5. Click on parent page
            //6. Click "Delete" link
            mainPage.ClickDeleteLink(parentPage.PageName);

            //VP. Check confirm message "Are you sure you want to remove this page?" appears
            //7. Click OK button

            string actualMsg = mainPage.GetDialogText();
            string expectedMsg = "Are you sure you want to remove this page?";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
            mainPage.ConfirmDialog("OK");

            //VP. Check warning message "Can not delete page 'Test' since it has children page(s)" appears
            //8. Click OK button
            actualMsg = mainPage.GetDialogText();
            expectedMsg = "Cannot delete page '" + parentPage.PageName + "' since it has child page(s).\r\n";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
            mainPage.ConfirmDialog("OK");

            //9. Click on  children page
            //10. Click "Delete" link
            mainPage.ClickDeleteLink(parentPage.PageName + "/" + childPage.PageName);

            //VP. Check confirm message "Are you sure you want to remove this page?" appears
            //11. Click OK button
            actualMsg = mainPage.GetDialogText();
            expectedMsg = "Are you sure you want to remove this page?";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
            mainPage.ConfirmDialog("OK");
            mainPage.WaitForPageLoadComplete();

            //VP. Check children page is deleted
            Assert.IsFalse(mainPage.IsPageExist(parentPage.PageName + "/" + childPage.PageName), childPage.PageName + " is still exist");

            //12. Click on  parent page
            //13. Click "Delete" link
            mainPage.ClickDeleteLink(parentPage.PageName);

            //VP. Check confirm message "Are you sure you want to remove this page?" appears
            //14. Click OK button
            actualMsg = mainPage.GetDialogText();
            expectedMsg = "Are you sure you want to remove this page?";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
            mainPage.ConfirmDialog("OK");
            mainPage.WaitForPageLoadComplete();

            //VP. Check parent page is deleted
            Assert.IsFalse(mainPage.IsPageExist(parentPage.PageName), parentPage.PageName + " is still exist");

            //15. Click on "Overview" page
            //VP. Check "Delete" link disappears
            mainPage.GoToPage("Overview");
            mainPage.LblGlobalSetting.MouseOver();
            Assert.IsFalse(mainPage.IsLinkExist("Delete"));
        }

        /// <summary>
        /// Verify that user is able to add additional sibbling pages to the parent page successfully
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC018()
        {
            Console.WriteLine("DA_MP_TC018 - Verify that user is able to add additional sibbling pages to the parent page successfully");

            //Set variables
            Page parentPage = new Page(CommonAction.GeneratePageName(), "Select parent", 2, "Overview", false);
            Page childPage1 = new Page(CommonAction.GeneratePageName(), parentPage.PageName, 3, "Select page", false);
            Page childPage2 = new Page(CommonAction.GeneratePageName(), parentPage.PageName, 2, "Select page", false);

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Log in specific repository with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter Page Name (Test)
            //5. Click OK button
            //6. Go to Global Setting -> Add page
            //7. Enter Page Name (Test Child 1)
            //8. Click on  Parent Page dropdown list
            //9. Select a parent page (Test)
            //10. Click OK button
            //11. Go to Global Setting -> Add page
            //12. Enter Page Name (Test Child 2)
            //13. Click on  Parent Page dropdown list
            //14. Select a parent page (Test)
            //15. Click OK button
            mainPage.AddPage(parentPage);
            mainPage.AddPage(childPage1);
            mainPage.AddPage(childPage2);

            //VP. Check "Test Child 2" is added successfully
            Assert.IsTrue(mainPage.IsPageExist(parentPage.PageName + "/" + childPage2.PageName), string.Format("{0} is not exist", childPage2.PageName));

            //Post-Condition
            mainPage.DeletePage(parentPage.PageName + "/" + childPage1.PageName);
            mainPage.DeletePage(parentPage.PageName + "/" + childPage2.PageName);
            mainPage.DeletePage(parentPage.PageName);
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that user is able to add additional sibbling page levels to the parent page successfully
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC019()
        {
            Console.WriteLine("DA_MP_TC019 - Verify that user is able to add additional sibbling page levels to the parent page successfully.");

            //Set variables
            Page page = new Page(CommonAction.GeneratePageName(), "Overview", 2, "Select page", false);

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog (Page name: Page 1, Parent page: Overview)
            //5. Click OK button
            mainPage.AddPage(page);

            //VP. Observe the current page
            //    User is able to add additional sibbling page levels to parent page successfully. 
            //    In this case: Overview is parent page of page 1.
            mainPage.GoToLink("Overview");
            Assert.IsTrue(mainPage.IsPageExist("Overview/" + page.PageName), string.Format("{0} is not exist", page.PageName));

            //Post-Condition
            mainPage.DeletePage("Overview/" + page.PageName);
        }

        /// <summary>
        /// Verify that user is able to delete sibbling page as long as that page has not children page under it
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC020()
        {
            Console.WriteLine("DA_MP_TC020 - Verify that user is able to delete sibbling page as long as that page has not children page under it");
            //Set variables
            Page page1 = new Page(CommonAction.GeneratePageName(), "Overview", 2, "Select page", false);
            Page page2 = new Page(CommonAction.GeneratePageName(), "Overview/" + page1.PageName, 2, "Select page", false);

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog (Page name: Page 1, Parent page: Overview)
            //5. Go to Global Setting -> Add page
            //6. Enter info into all required fields on New Page dialog (Page name: Page 2, Parent page: Page 1)
            mainPage.AddPage(page1);
            mainPage.AddPage(page2);

            //7. Go to the first created page
            //8. Click Delete link
            //9. Click Ok button on Confirmation Delete page
            mainPage.ClickDeleteLink("Overview/" + page1.PageName);
            mainPage.ConfirmDialog("OK");

            //VP. Observe the current page
            //    There is a message "Can't delete page "page 1" since it has children page"
            //10. Close confirmation dialog
            string actualMsg = mainPage.GetDialogText();
            string expectedMsg = string.Format("Cannot delete page '{0}' since it has child page(s).\r\n", page1.PageName);
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect: {0}", actualMsg));
            mainPage.ConfirmDialog("OK");

            //11. Go to the second page
            //12. Click Delete link
            //13. Click Ok button on Confirmation Delete page
            mainPage.DeletePage("Overview/" + page1.PageName + "/" + page2.PageName);

            //VP. Observe the current page
            //    Page 2 is deleted successfully
            mainPage.GoToLink("Overview/" + page1.PageName);
            Assert.IsFalse(mainPage.IsPageExist("Overview/" + page1.PageName + "/" + page2.PageName), string.Format("{0} is still exist!", page2.PageName));

            //Post-Condition
            mainPage.DeletePage("Overview/" + page1.PageName);
        }

        /// <summary>
        /// Verify that user is able to edit the name of the page (Parent/Sibbling) successfully
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC021()
        {
            Console.WriteLine("DA_MP_TC021 - Verify that user is able to edit the name of the page (Parent/Sibbling) successfully");
            Page page1 = new Page(CommonAction.GeneratePageName(), "Overview", 2, "Select page", false);
            Page page2 = new Page(CommonAction.GeneratePageName(), "Overview/" + page1.PageName, 2, "Select page", false);

            //1. Navigate to Dashboard login page
            //2. Login with valid account
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog (Page name: Page 1, Parent page: Overview)
            //5. Go to Global Setting -> Add page
            //6. Enter info into all required fields on New Page dialog (Page name: Page 2, Parent page: Page 1)
            mainPage.AddPage(page1);
            mainPage.AddPage(page2);

            //7. Go to the first created page
            //8. Click Edit link
            //9. Enter another name into Page Name field (Page name: Page 3)
            //10. Click Ok button on Edit Page dialog
            mainPage.OpenEditPage(page1.ParentPage + "/" + page1.PageName);
            page1.PageName = CommonAction.GeneratePageName();
            mainPage.EditPage(page1);

            //VP. Observe the current page (User is able to edit the name of parent page successfully)
            Assert.IsTrue(mainPage.IsPageExist("Overview/" + page1.PageName), string.Format("{0} is not exist", page1.PageName));

            //11. Go to the second created page
            //12. Click Edit link
            //13. Enter another name into Page Name field
            //14. Click Ok button on Edit Page dialog
            mainPage.OpenEditPage("Overview/" + page1.PageName + "/" + page2.PageName);
            page2.PageName = CommonAction.GeneratePageName();
            page2.ParentPage = "Overview/" + page1.PageName;
            mainPage.EditPage(page2);

            //VP. Observe the current page (User is able to edit the name of sibbling page successfully)
            Assert.IsTrue(mainPage.IsPageExist("Overview/" + page1.PageName + "/" + page2.PageName), string.Format("{0} is not exist", page2.PageName));

            //Post-Condition
            mainPage.DeletePage("Overview/" + page1.PageName + "/" + page2.PageName);
            mainPage.DeletePage("Overview/" + page1.PageName);
        }

        /// <summary>
        /// Verify that user is unable to duplicate the name of sibbling page under the same parent page
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC022()
        {
            Console.WriteLine("DA_MP_TC022 - Verify that user is unable to duplicate the name of sibbling page under the same parent page");

            //Set Variable
            string pageName = CommonAction.GeneratePageName();
            string childPageName = "Child" + CommonAction.GeneratePageName();

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Add a new page
            //4. Add a sibling page of new page
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            Page pageSibling = new Page(childPageName, pageName, 2, "Select page", false);
            mainPage.AddPage(page).AddPage(pageSibling);

            //5. Go to Global Setting -> Add page
            //6. Enter Page Name
            //7. Click on  Parent Page dropdown list
            //8. Select a parent page
            //9. Click OK button
            string actualMessage = mainPage.AddPageWithError(pageSibling);

            //VP. Check warning message "Test child already exist. Please enter a diffrerent name" appears
            Assert.IsTrue(actualMessage.Contains(childPageName + " already exists. Please enter a different name."), string.Format("Message incorrect {0}", actualMessage));

            //Close message and close add new page dialog
            mainPage.ConfirmDialog("OK");
            AddNewPage newPage = new AddNewPage(_webDriver);
            newPage.CancelPage();

            //Post-Condition
            mainPage.DeletePage(page.PageName + "/" + pageSibling.PageName);
            mainPage.DeletePage(page.PageName);
            mainPage.Logout();

        }

        /// <summary>
        /// Verify that user is able to edit the parent page of the sibbling page successfully
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC023()
        {
            Console.WriteLine("DA_MP_TC023 - Verify that user is able to edit the parent page of the sibbling page successfully");

            //Set Variable
            string pageName = CommonAction.GeneratePageName();
            string childPageName = "Child" + CommonAction.GeneratePageName();
            string pageName2 = CommonAction.GeneratePageName();

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog
            //5. Go to Global Setting -> Add page
            //6. Enter info into all required fields on New Page dialog
            //7. Go to the first created page
            //8. Click Edit link
            //9. Enter another name into Page Name field
            //10. Click Ok button on Edit Page dialog
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            Page pageSibling = new Page(childPageName, pageName, 2, "Select page", false);
            Page pageEdit = new Page(pageName2, null, 2, null, false);
            mainPage.AddPage(page).AddPage(pageSibling);
            mainPage.OpenEditPage(pageName).EditPage(pageName, pageEdit);

            //VP: User is able to edit the parent page of the sibbling page successfully
            Assert.IsTrue(mainPage.IsPageExisted(pageName2), "User can edit the parent page of sibling page");

            //Post-Condition

            mainPage.DeletePage(pageName2 + "/" + pageSibling.PageName);
            mainPage.DeletePage(pageName2);
            mainPage.Logout();

        }

        /// <summary>
        /// Verify that "Bread Crums" navigation is correct
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC024()
        {
            Console.WriteLine("DA_MP_TC024 - Verify that \"Bread Crums\" navigation is correct");

            //Set Variable
            string pageName = CommonAction.GeneratePageName();
            string childPageName = "Child" + CommonAction.GeneratePageName();

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog
            Page page = new Page(pageName, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page);

            string page1URL = mainPage.GetURL();

            //5. Go to Global Setting -> Add page
            //6. Enter info into all required fields on New Page dialog
            Page pageSibling = new Page(childPageName, pageName, 2, "Select page", false);
            mainPage.AddPage(pageSibling);

            string page2URL = mainPage.GetURL();

            //5. Click the first breadcrums (Page1)
            //VP: The first page is navigated
            mainPage.GoToPage(page.PageName);
            string actualPage1 = mainPage.GetURL();
            Assert.AreEqual(page1URL, actualPage1, "The first page is navigated");

            //6. Click the second breadcrums
            //VP: The second page is navigated
            mainPage.GoToPage(pageSibling.PageName);
            string actualPage2 = mainPage.GetURL();
            Assert.AreEqual(page2URL, actualPage2, "The second page is navigated");

            //Post-Condition
            mainPage.DeletePage(page.PageName + "/" + pageSibling.PageName);
            mainPage.DeletePage(page.PageName);
            mainPage.Logout();

        }

        /// <summary>
        /// Verify that page listing is correct when user edit "Display After" field of a specific page
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC025()
        {
            Console.WriteLine("DA_MP_TC025 - Verify that page listing is correct when user edit \"Display After\" field of a specific page");

            //Set Variable
            string pageName1 = CommonAction.GeneratePageName();
            string pageName2 = CommonAction.GeneratePageName();

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog
            //5. Go to Global Setting -> Add page
            //6. Enter info into all required fields on New Page dialog
            Page page = new Page(pageName1, "Select parent", 2, "Overview", false);
            Page page2 = new Page(pageName2, "Select parent", 2, "Overview", false);
            mainPage.AddPage(page).AddPage(page2);

            //5. Click Edit link for the second created page
            mainPage.GoToPage(pageName2).OpenSetting().LnkEditPage.Click();

            //6. Change value Display After for the second created page to after Overview page
            //7. Click Ok button on Edit Page dialog
            Page pageEdit = new Page(pageName2, null, 2, "Overview", false);
            mainPage.EditPage(pageEdit);

            //VP: Position of the second page follow Overview page
            bool actualResult = mainPage.GetPositionPage(pageName2) > mainPage.GetPositionPage("Overview") ? true : false;
            Assert.AreEqual(true, actualResult);

            //Post-Condition
            mainPage.DeletePage(pageName1);
            mainPage.DeletePage(pageName2);
            mainPage.Logout();

        }

        /// <summary>
        /// Verify that page column is correct when user edit "Number of Columns" field of a specific page
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_MP_TC026()
        {
            Console.WriteLine("DA_MP_TC026 - Verify that page column is correct when user edit \"Number of Columns\" field of a specific page");

            //Set Variable
            string pageName1 = CommonAction.GeneratePageName();

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver);
            loginPage.Open();

            //2. Login with valid account
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Go to Global Setting -> Add page
            //4. Enter info into all required fields on New Page dialog
            //Page name: Page 1; Number of Columns: 2
            Page page = new Page(pageName1, "Select parent", 2, "Overview", false);
            Page page2 = new Page(pageName1, null, 3, null, false);
            mainPage.AddPage(page);
            mainPage.OpenSetting().LnkEditPage.Click();

            //5. Go to Global Setting -> Edit link
            //6. Edit Number of Columns for the above created page (Number of Columns: 3)
            //7. Click OK button
            mainPage.EditPage(page2);

            //VP. There are 3 columns on the above created page
            int actualColumnNumber = mainPage.GetColumnCount();
            Assert.AreEqual(3, actualColumnNumber, "There are " + actualColumnNumber + " column");

            //Post-Condition
            mainPage.DeletePage(page.PageName);
            mainPage.Logout();

        }
    }
}
