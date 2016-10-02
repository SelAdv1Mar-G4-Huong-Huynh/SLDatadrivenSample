using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumAdvProject.PageObjects;
using SeleniumAdvProject.Common;
using SeleniumAdvProject.TestDataAccess;

namespace SeleniumAdvProject.TestCases
{
    [TestClass]
    public class LoginTestCases : BaseTestCase
    {
        /// <summary>
        /// DA_LOGIN_TC001 - Verify that user can login specific repository successfully via Dashboard login page with correct credentials
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC001()
        {
            Console.WriteLine("DA_LOGIN_TC001 - Verify that user can login specific repository successfully via Dashboard login page with correct credentials");
            
            //1. Navigate to Dashboard login page
            //2. Enter valid username and password	
            //3. Click on "Login" button
            LoginPage loginPage = new LoginPage(_webDriver).Open();            
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //VP. Verify that Dashboard Mainpage appears
            Assert.AreEqual(Constants.UserName, mainPage.GetUserNameText(), "Login is not successfully");

            //Post-Condition
            mainPage.Logout();
        }
        [TestMethod]
        public void DA_LOGIN_TC_DATA_DRIVEN_DEMO001()
        {
            Console.WriteLine("DA_LOGIN_TC001 - Verify that user can login specific repository successfully via Dashboard login page with correct credentials");
            ExcelLib.PopulateInCollection(@"D:\Logigear\Selenium csharp\C# Selenium\Advance\Homewok\G4AdvSeleniumDataDriven\SeleniumAdvProject\SeleniumAdvProject\TestDataAccess\TestData.xlsx");

            //1. Navigate to Dashboard login page
            //2. Enter valid username and password	
            //3. Click on "Login" button
            LoginPage loginPage = new LoginPage(_webDriver).Open();
            MainPage mainPage = loginPage.Login(ExcelLib.ReadData(1, "Repository"), ExcelLib.ReadData(1, "UserName"), ExcelLib.ReadData(1, "Password"));

            //VP. Verify that Dashboard Mainpage appears
            Assert.AreEqual(Constants.UserName, mainPage.GetUserNameText(), "Login is not successfully");

            //Post-Condition
            mainPage.Logout();
        }
        /// <summary>
        /// Verify that user fails to login specific repository successfully via Dashboard login page with incorrect credentials
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC002()
        {
            Console.WriteLine("DA_LOGIN_TC002 - Verify that user fails to login specific repository successfully via Dashboard login page with incorrect credentials");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Enter invalid username and password
            //3. Click on "Login" button
            string actualMsg = loginPage.LoginWithExpectedError(Constants.Repository, "abc", "abc");

            //4. Verify that Dashboard Error message "Username or password is invalid" appears
            string expectedMsg = "Username or password is invalid";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
        }

        /// <summary>
        /// Verify that user fails to log in specific repository successfully via Dashboard login page with correct username and incorrect password
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC003()
        {
            Console.WriteLine("DA_LOGIN_TC003 - Verify that user fails to log in specific repository successfully via Dashboard login page with correct username and incorrect password");

            //1 Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2 Enter valid username and invalid password (administrator / abc)
            //3 Click on "Login" button
            string actualMsg = loginPage.LoginWithExpectedError(Constants.Repository, Constants.UserName, "abc");

            //4 Verify that Dashboard Error message "Username or password is invalid" appears
            string expectedMsg = "Username or password is invalid";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
        }

        /// <summary>
        /// Verify that user is able to log in different repositories successfully after logging out current repository
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC004()
        {
            Console.WriteLine("DA_LOGIN_TC004 - Verify that user is able to log in different repositories successfully after logging out current repository");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Enter valid username and password of default repository (administrator / <blank>)
            //3. Click on "Login" button
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //4. Click on "Logout" button
            mainPage.Logout();

            //5. Select a different repository (TestRepository)
            //6. Enter valid username and password of this repository (administrator / <blank>)
            loginPage.Login(Constants.TestRepository, Constants.UserName, Constants.Password);

            //VP. Verify that Dashboard Mainpage appears
            Assert.IsTrue(mainPage.Displayed(Constants.UserName), "Main page does not appear");

            //Post-Condition
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that there is no Login dialog when switching between 2 repositories with the same account
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC005()
        {
            Console.WriteLine("DA_LOGIN_TC005 - Verify that there is no Login dialog when switching between 2 repositories with the same account");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with valid account for the first repository
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);

            //3. Choose another repository in Repository list (TestRepository)
            mainPage.SelectRepository(Constants.TestRepository);

            //VP Observe the current page
            //- There is no Login Repository dialog
            //- The Repository menu displays name of switched repository
            Assert.IsFalse(loginPage.Displayed(), "Login Repository dialog is displaying");
            Assert.AreEqual(Constants.TestRepository, mainPage.GetCurrentRepositoryText(), "The Repository menu displays name switched repository is not exist");
        }

        /// <summary>
        /// Verify that "Password" input is case sensitive
        /// </summary>
        /// <author>Vu Tran</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC006()
        {
            Console.WriteLine("DA_LOGIN_TC006 - Verify that \"Password\" input is case sensitive");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with the account has uppercase password
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName2, Constants.Password2);

            //VP. Observe the current page
            // Main page is displayed
            Assert.IsTrue(mainPage.Displayed(Constants.UserName2), "Main page is NOT displaying");

            //3. Logout TA Dashboard
            mainPage.Logout();

            //4. Login with the above account but enter lowercase password
            string actualMsg = loginPage.LoginWithExpectedError(Constants.Repository, Constants.UserName2, Constants.Password2.ToLower());

            //VP. Dashboard Error message "Username or password is invalid" appears
            string expectedMsg = "Username or password is invalid";
            Assert.AreEqual(expectedMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
        }

        /// <summary>
        /// Verify that "Username" input is not case sensitive
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC007()
        {
            Console.WriteLine("DA_LOGIN_TC007 - Verify that \"Username\" input is not case sensitive");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with the account has uppercase username
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UpperCaseUserName, Constants.Password);

            //VP. Observe the current page
            // Main page is displayed
            bool isDisplay = mainPage.Displayed(Constants.UpperCaseUserName);
            Assert.AreEqual(true, isDisplay, "User login unsuccessfully with uppercase username");
            
            //3. Logout TA Dashboard
            mainPage.Logout();

            //4. Login with the above account but enter lowercase username
            //VP. Main page is displayed
            loginPage.Login(Constants.Repository, Constants.LowerCaseUser, Constants.Password);
            Assert.IsTrue(mainPage.Displayed(Constants.LowerCaseUser.ToUpper()), "User login unsuccessfully with uppercase username");
        }

        /// <summary>
        /// Verify that password with special characters is working correctly
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC008()
        {
            Console.WriteLine("DA_LOGIN_TC008 - Verify that password with special characters is working correctly");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with account that has special characters password
            //VP. Main page is displayed
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName1, Constants.SpecialPassword);
            Assert.IsTrue(mainPage.Displayed(Constants.UserName1), "User login unsuccessfully with special characters password");

            //Post-Condition
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that username with special characters is working correctly
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC009()
        {
            Console.WriteLine("DA_LOGIN_TC009 - Verify that username with special characters is working correctly");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Login with account that has special characters username
            //VP. Main page is displayed
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.SpecialUserName, Constants.Password);
            Assert.IsTrue(mainPage.Displayed(Constants.SpecialUserName), "User login unsuccessfully with special characters username");

            //Post-Condition
            mainPage.Logout();
        }

        /// <summary>
        /// Verify that the page works correctly for the case when no input entered to Password and Username field
        /// </summary>
        /// <author>Tu Nguyen</author>
        /// <date>05/25/2015</date>
        [TestMethod]
        public void DA_LOGIN_TC010()
        {
            Console.WriteLine("DA_LOGIN_TC010 - Verify that the page works correctly for the case when no input entered to Password and Username field");

            //1 Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Click Login button without entering data into Username and Password field
            string actualMsg = loginPage.LoginWithOutAccount(Constants.Repository);

            //VP. Observe the current page
            // There is a message "Please enter username"
            string expectMsg = "Please enter username";
            Assert.AreEqual(expectMsg, actualMsg, string.Format("Message incorrect {0}", actualMsg));
        }
    }
}
