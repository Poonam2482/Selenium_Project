using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Configuration;
using System.Collections.Specialized;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;

namespace ConsoleApp1
{
    [TestClass]
    class Login
    {
        IWebDriver m_driver;
        DataTable dtLog = new DataTable();
        DataSet ds = new DataSet();
        #region "Login Automation"
        public void Login_Automation()
        {
            string TestID = "", SrNo = "", Module = "", ExpectedResult = "", AlertText = "Pass";
            try
            {
                string filePath = ConfigurationManager.AppSettings["InputFilePath"];
                string fileExt = Path.GetExtension(filePath); //get the file extension
                ReadExcel(filePath, fileExt);

                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    int TotalTest = ds.Tables[0].Rows.Count;
                    string URL = ds.Tables[0].Rows[0]["URL"].ToString();
                    m_driver = new ChromeDriver(ConfigurationManager.AppSettings["DriverPath"]);
                    //m_driver = new InternetExplorerDriver(ConfigurationManager.AppSettings["DriverPath"]);
                    m_driver.Url = URL;
                    m_driver.Manage().Window.Maximize();

                    for (int i = 0; i < TotalTest; i++)
                    {
                        Module = ds.Tables[0].Rows[i]["Module"].ToString();
                        TestID = ds.Tables[0].Rows[i]["TestID"].ToString();

                        string Action = ds.Tables[0].Rows[i]["Action"].ToString();
                        int Columns = Convert.ToInt32(ds.Tables[1].Rows[0]["Columns"].ToString());

                        string id, FieldType, value, alert, Result = "";

                        DataRow[] Table_inner = ds.Tables[1].Select("TestId = '" + Convert.ToString(TestID) + "'");
                        if (Module == "Menu")
                        {
                            try
                            {

                                Actions objActions = new Actions(m_driver);
                                string SubMenuText = Table_inner[0]["SubModule"].ToString();
                                string SubModule = Table_inner[0]["SubSubModule"].ToString();
                                string MenuURL = Table_inner[0]["URL"].ToString();

                                Thread.Sleep(1000);

                                ReadOnlyCollection<IWebElement> menuItemList = m_driver.FindElements(By.TagName("a"));
                                foreach (var item in menuItemList)
                                {
                                    if (item.Text.ToLower().Equals(SubMenuText.ToLower()))
                                    {
                                        var Item1 = item;
                                        objActions.MoveToElement(Item1).Perform();
                                        Thread.Sleep(2000);
                                        if (SubModule == "")
                                        {
                                            m_driver.FindElement(By.PartialLinkText(Item1.Text)).Click();
                                        }
                                        else
                                        {
                                            m_driver.FindElement(By.PartialLinkText(SubModule)).Click();
                                            Thread.Sleep(2000);
                                            m_driver.FindElement(By.PartialLinkText(MenuURL)).Click();
                                        }
                                        break;
                                    }
                                }
                                if (Action.Trim() != "")
                                {
                                    m_driver.FindElement(By.Id(Action)).Click();
                                    Thread.Sleep(1000);
                                }
                                InsertLogRow(TestID, SrNo, "Pass", "", Module, AlertText, ExpectedResult);
                                Thread.Sleep(1000);
                            }
                            catch (Exception e)
                            {
                                InsertLogRow(TestID, SrNo, "Fail", e.ToString(), Module, "", "");
                            }
                        }
                        else if (Module == "Login" || Module == "Form")
                        {
                            for (var x = 0; x < Table_inner.Length; x++)
                            {
                                for (int j = 1; j <= Columns; j++)
                                {
                                    SrNo = Table_inner[x]["SrNo"].ToString();
                                    id = ds.Tables[0].Rows[i]["Field" + j].ToString();
                                    FieldType = ds.Tables[0].Rows[i]["FieldType" + j].ToString();
                                    value = Table_inner[x]["field" + j].ToString();
                                    Result = Table_inner[x]["OutPutLabel"].ToString();

                                    if (FieldType == "TextBox")
                                    {
                                        m_driver.FindElement(By.CssSelector("input[id=" + id + "]")).Clear();
                                        m_driver.FindElement(By.CssSelector("input[id=" + id + "]")).SendKeys(value);
                                    }
                                    if (FieldType == "DropDownList")
                                    {
                                        string LocationValue = value;
                                        if (value != "")
                                        {
                                            SelectElement Location = new SelectElement(m_driver.FindElement(By.CssSelector("Select[id=" + id + "]")));
                                            Location.SelectByValue(value);
                                        }
                                        string abc = "cyx";
                                    }
                                    //Thread.Sleep(5000);
                                }
                                if (Action == "ctl00_ContentPlaceHolder1_btncheck")
                                {
                                }
                                IWebElement btnsubmit = m_driver.FindElement(By.CssSelector("input[id=" + Action + "]"));
                                btnsubmit.Click();
                                Thread.Sleep(2000);
                                alert = Table_inner[x]["Alert"].ToString();
                                ExpectedResult = "";
                                ExpectedResult = Table_inner[x]["ExpectedResult"].ToString();
                                AlertText = "";
                                if (alert == "Yes")
                                {
                                    IAlert AlertMsg = m_driver.SwitchTo().Alert();
                                    AlertText = AlertMsg.Text;
                                    Console.WriteLine("Alert Text Is : " + AlertText);

                                    Thread.Sleep(5000);
                                    AlertMsg.Accept();

                                    Assert.AreEqual(ExpectedResult, AlertText);
                                    InsertLogRow(TestID, SrNo, "Pass", "", Module, AlertText, ExpectedResult);

                                }
                                else if (alert == "No")
                                {
                                    IWebElement msg = null;
                                    if (Result != null && Result != "")
                                    {
                                        try
                                        {
                                            msg = m_driver.FindElement(By.CssSelector("span[id=" + Result + "]"));
                                            if (msg != null)
                                            {
                                                Assert.AreEqual(ExpectedResult, msg.Text);
                                                InsertLogRow(TestID, SrNo, "Pass", "", Module, msg.Text, ExpectedResult);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            InsertLogRow(TestID, SrNo, "Fail", "", Module, msg.Text, ExpectedResult);
                                        }
                                    }
                                    else
                                    {
                                        InsertLogRow(TestID, SrNo, "Pass", "", Module, "", ExpectedResult);
                                    }
                                }
                                else if (alert == "Link")
                                {
                                    IWebElement msg = null;
                                    if (Result != null && Result != "")
                                    {
                                        try
                                        {
                                            msg = m_driver.FindElement(By.CssSelector("a[id=" + Result + "]"));
                                            if (msg != null)
                                            {
                                                Assert.AreEqual(ExpectedResult, msg.Text);
                                                InsertLogRow(TestID, SrNo, "Pass", "", Module, msg.Text, ExpectedResult);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            InsertLogRow(TestID, SrNo, "Fail", "", Module, msg.Text, ExpectedResult);
                                        }
                                    }
                                    else
                                    {
                                        InsertLogRow(TestID, SrNo, "Pass", "", Module, "", ExpectedResult);
                                    }
                                }
                                else
                                {
                                    InsertLogRow(TestID, SrNo, "Pass", "", Module, AlertText, ExpectedResult);
                                    //Thread.Sleep(5000);
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                IWebElement btnsubmit = m_driver.FindElement(By.CssSelector("a[id=" + Action + "]"));
                                btnsubmit.Click();
                                InsertLogRow(TestID, SrNo, "Pass", "", Module, AlertText, "Pass");
                            }
                            catch (Exception ex)
                            {
                                InsertLogRow(TestID, SrNo, "Fail", ex.ToString(), Module, AlertText, "Fail");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                InsertLogRow(TestID, SrNo, "Fail", ex.ToString(), Module, AlertText, ExpectedResult);
                //m_driver.Quit();
            }
            finally
            {
                var outputFileName = ConfigurationManager.AppSettings["OutputPath"] + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";

                FileInfo objFileInfo = new FileInfo(outputFileName);

                using (ExcelPackage pck = new ExcelPackage(objFileInfo))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Test Cases");
                    ws.Cells["A1"].LoadFromDataTable(dtLog, true);
                    pck.Save();
                }
                //m_driver.Quit();
            }
        }
        #endregion

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(ds, "Table1"); //fill excel data into dataTable 

                    oleAdpt = new OleDbDataAdapter("select * from [Sheet2$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(ds, "Table2"); //fill excel data into dataTable  
                }
                catch (Exception ex)
                { }
            }
            return dtexcel;
        }

        public void InsertLogRow(string testId, string srNo, string result, object strMessage, string Module, string ActualResult, string ExpectedResult)
        {
            string strReason = "";

            if (dtLog.Columns.Count < 1)
            {
                dtLog.Columns.Add("Test Id");
                dtLog.Columns.Add("Module");
                dtLog.Columns.Add("Sr No");
                dtLog.Columns.Add("ActualResult");
                dtLog.Columns.Add("ExpectedResult");
                dtLog.Columns.Add("Test Result");
                dtLog.Columns.Add("Test Message");
            }

            dtLog.NewRow();
            if (strMessage == null)
            {
                strReason = "";
            }
            else
            {
                strReason = strMessage.ToString();
            }
            dtLog.Rows.Add(testId, Module, srNo, ActualResult, ExpectedResult, result, strReason);
        }

    }

}
