using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.Ultilities.Controls
{
    public class Table:BaseControl
    {  
        public Table() { }
        public Table(string xPath) : base(By.XPath(xPath)) { }
        public Table(IWebElement element) : base(element) { }
        public Table(By by) : base(by) { }
        
        public Table(IWebElement element,IWebDriver webDriver){}

        public int RowsCount()
        {
            return FindElements(By.TagName("tr")).Count;
        }

        /// <summary>
        /// Gets all rows in table
        /// </summary>
        public List<Tr> Rows
        {
            get
            {
                List<Tr> lst = new List<Tr>();
                IList<IWebElement> list = this.FindElements(By.TagName("tr"));
                for (int i = 1; i <= list.Count; i++)
                {
                    lst.Add(new Tr(String.Format("{0}//tr[{1}]",this.XPath, i)));
                }

                return lst;
            }
        }

        /// <summary>
        /// Count column index in table
        /// </summary>
        public int ColumnCount()
        {
            List<Tr> row = new List<Tr>();
            row = Rows;
            return row[0].Cells.Count;
        }        
       
        public List<String> GetTableTextContent
        {
            get
            {
                List<string> texts = new List<string>();
                List<Tr> trs = Rows;
                foreach(Tr tr in trs){
                    List<Td> tds = tr.Cells;
                    foreach (Td td in tds)
                    {
                        texts.Add(td.Text);
                    }
                }

                return texts;
            }
        }
    }
}
