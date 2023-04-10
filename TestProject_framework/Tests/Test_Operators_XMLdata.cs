
using NPOI.XWPF.UserModel;
using OpenQA.Selenium;
using System;
using System.Data;
using System.Linq;
using System.Reflection.PortableExecutable;
using TestProject_framework.Helpers;
using TestProject_framework.Reader;

namespace TestProject_framework.Tests
{
    public class Test_Operators_XMLdata : DriverHelper
    {
        xmlreaderclass cn = new xmlreaderclass();
        [OneTimeSetUp]
        public void Setup()
        {


            string ValueofName = cn.DataReader("URL");
            if (ValueofName != null)
            {
                BrowserHelpers b = new BrowserHelpers();
                driver = b.ChromeInitialize(driver, ValueofName);

            }



        }
        [Test, Order(0), Category("Seraching All operators in home page")]
        public void SearchOperator()
        {
            GenericHelpers.CurrentTestName();
            JSExecutorHelper js = new JSExecutorHelper();
            GenericHelpers gp = new GenericHelpers();
            try
            {

                string ValueofName = cn.DataReader("Homeoperator");
                if (ValueofName != null)
                {
                    var ele = driver.FindElement(By.XPath(ValueofName));
                    JSExecutorHelper.jsjs(driver, ele);
                    Thread.Sleep(1000);
                    GenericHelpers.Clickelement(driver, cn.DataReader("Homeoperator"));
                    GenericHelpers.switchtabs(driver);
                }
            }
            catch (Exception ex)
            {
                if (driver != null)
                    driver.Quit();
                throw;
            }


        }
        [Test, Order(1), Category("Search I operators")]
        public void SearchWithIndexOperator_I()
        {

            GenericHelpers.CurrentTestName();
            try
            {

                string dir = cn.DataReader("Directory");
                string page = cn.DataReader("operatorPage");
                if (dir != null)
                {
                    string val = dir + "i";
                    if (page != null)
                    {
                        GenericHelpers.searchElement(driver, page, val);
                        driver.SwitchTo().ActiveElement();
                    }
                }
            }
            catch (Exception ex) { throw; }


        }
        [Test, Order(2), Category("Get pagination for I operator")]
        public void pagination_I()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("I_operator_pages") != null)
                    GenericHelpers.searchPages(driver, cn.DataReader("I_operator_pages"), "3");
            }
            catch (Exception ex) { throw; }
        }
        [Test, Order(3), Category("Print all bus operators for I operator")]
        public void PrintAllRecordsforPage_I()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("I_operator") != null)
                    GenericHelpers.readallelements(driver, cn.DataReader("I_operator"), "I");
            }
            catch (Exception ex) { throw; }

        }
        [Test, Order(4), Category("Search R operators")]
        public void SearchWithIndexOperator_R()
        {

            GenericHelpers.CurrentTestName();
            try
            {
                string dir = cn.DataReader("Directory");
                string page = cn.DataReader("operatorPage");
                if (dir != null)
                {
                    string val = dir + "r";
                    if (page != null)
                    {
                        GenericHelpers.searchElement(driver, page, val);
                        driver.SwitchTo().ActiveElement();
                    }
                }
            }
            catch (Exception ex) { throw; }

        }
        [Test, Order(5), Category("Get pagination for R operator")]
        public void pagination_R()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("R_operator_pages") != null)
                    GenericHelpers.searchPages(driver, cn.DataReader("R_operator_pages"), "3");
            }
            catch (Exception ex) { throw; }
        }
        [Test, Order(6), Category("Print all bus operators for R operator")]
        public void PrintAllRecordsforPage_R()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("R_operator") != null)
                    GenericHelpers.readallelements(driver, cn.DataReader("R_operator"), "R");
            }
            catch (Exception ex) { throw; }

        }
        [Test, Order(7), Category("Search A operators")]
        public void SearchWithIndexOperator_A()
        {
            GenericHelpers.CurrentTestName();

            try
            {
                string dir = cn.DataReader("Directory");
                string page = cn.DataReader("operatorPage");
                if (dir != null)
                {
                    string val = dir + "a";
                    if (page != null)
                    {
                        GenericHelpers.searchElement(driver, page, val);
                        driver.SwitchTo().ActiveElement();
                    }
                }
            }
            catch (Exception ex) { throw; }

        }

        [Test, Order(8), Category("Get pagination for A operator")]
        public void pagination_A()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("A_operator_pages") != null)
                    GenericHelpers.searchPages(driver, cn.DataReader("A_operator_pages"), "3");
            }
            catch (Exception ex) { throw; }
        }
        [Test, Order(9), Category("Print all bus operators for A operator")]
        public void PrintAllRecordsforPage_A()
        {
            GenericHelpers.CurrentTestName();
            try
            {
                if (cn.DataReader("A_operator") != null)
                    GenericHelpers.readallelements(driver, cn.DataReader("A_operator"), "A");
            }
            catch (Exception ex) { throw; }

        }


        [OneTimeTearDown]
        public void oneTimeTeardown()
        {
            GenericHelpers.CurrentTestName();
            driver.Quit();
        }
    }
}