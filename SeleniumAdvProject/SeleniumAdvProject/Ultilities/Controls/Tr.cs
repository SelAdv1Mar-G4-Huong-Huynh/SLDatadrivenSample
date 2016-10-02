using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Tr: BaseControl
    {
        public Tr() { }
        public Tr(string xPath) : base(By.XPath(xPath)) { }
        public Tr(IWebElement element) : base(element) { }
        public Tr(By by) : base(by) { }

        /// <summary>
        /// Get all cells in table row
        /// </summary>
        public List<Td> Cells
        {
            get
            {
                List<Td> lst = new List<Td>();
                IList<IWebElement> list = FindElements(By.TagName("td"));
                for (int i = 1; i <= list.Count; i++)
                {
                    lst.Add(new Td(String.Format("{0}/td[{1}]", XPath, i)));
                }

                return lst;
            }
        }
    }
}
