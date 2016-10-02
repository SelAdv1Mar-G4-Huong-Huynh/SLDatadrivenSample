using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumAdvProject.PageObjects;
using SeleniumAdvProject.Common;

namespace SeleniumAdvProject.TestCases
{
    [TestClass]
    public class CleanUpTestCases : BaseTestCase
    {
        /// <summary>
        /// Clean created data after all all test suite
        /// </summary>
        /// <author>Huong Huynh</author>
        /// <date>06/06/2015</date>
        [TestMethod]
        public void DDeleteAll()
        {
            Console.WriteLine("DA_LOGIN_TC001 - Verify that user can login specific repository successfully via Dashboard login page with correct credentials");

            //1. Navigate to Dashboard login page
            LoginPage loginPage = new LoginPage(_webDriver).Open();

            //2. Enter valid username and password	
            //3. Click on "Login" button
            MainPage mainPage = loginPage.Login(Constants.Repository, Constants.UserName, Constants.Password);
            //mainPage.IsPageDisplayAfter
            // Delete all Page


            //Post-Condition
            mainPage.Logout();
        }        
    }
}
