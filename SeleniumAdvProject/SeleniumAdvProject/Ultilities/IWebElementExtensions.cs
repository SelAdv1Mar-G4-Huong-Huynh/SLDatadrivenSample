using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;

namespace SeleniumAdvProject.Ultilities
{
   public static class IWebElementExtensions
    {
       public static IWebElement MoveTo(this IWebElement element, IWebDriver webDriver)
       {
           Actions actions = new Actions(webDriver);
           actions.MoveToElement(element).Build().Perform();
           return element;
       }

       public static void Check(this IWebElement element)
       {
           if(!element.Selected){
               element.Click();
           }
       }

       public static void Uncheck(this IWebElement element)
       {
           if (element.Selected)
           {
               element.Click();
           }
       }
      
        
    }
}
