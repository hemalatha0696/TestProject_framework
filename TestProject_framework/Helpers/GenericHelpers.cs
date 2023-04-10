using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace TestProject_framework.Helpers
{
    public class GenericHelpers:DriverHelper
    {
        public static void Clickelement(IWebDriver driver, string element)
        {
            driver.FindElement(By.XPath(element)).Click();
            Thread.Sleep(1000);
        }
        public static void FindElement(IWebDriver driver, String element)
        {
            driver.FindElement(By.XPath(element));
            Thread.Sleep(1000);
          
        }
        public static void findallelements(IWebDriver driver, String[] allelement)
        {
            for (int all = 0; all < allelement.Length; all++)
            {
                driver.FindElements(By.XPath(allelement[all]));
               
            }
        }
        public static void searchElement(IWebDriver driver, String ele,String elementtobefound)
        {
            ReadOnlyCollection<IWebElement> alllist = driver.FindElements(By.XPath(ele));
            IList<IWebElement> allsublist = new List<IWebElement>(alllist);
            for (int i = 0; i < alllist.Count; i++)
            {
                String link = alllist[i].GetAttribute("href");
                if (link.Equals(elementtobefound))
                {
                    //JSExecutorHelper.jsjs(driver, alllist[i]);
                    alllist[i].Click();
                    break;
                }

                Thread.Sleep(1000);
            }
        }
        public static void searchPages(IWebDriver driver, String ele, String elementtobefound)
        {
            ReadOnlyCollection<IWebElement> alllist = driver.FindElements(By.XPath(ele));
            IList<IWebElement> allsublist = new List<IWebElement>(alllist);
            if (allsublist.Count == 1) JSExecutorHelper.totop(driver);
            for (int i = 0; i < alllist.Count; i++)
            {
                
                String link = alllist[i].Text;
                if (link.Equals(elementtobefound))
                {
                    JSExecutorHelper.jsjs(driver, alllist[i]);
                    alllist[i].Click();
                    break;
                }
                JSExecutorHelper.totop(driver);
                //Thread.Sleep(1000);
            }
        }
        public static void readallelements(IWebDriver driver, String ele,String path)
        {
            ReadOnlyCollection<IWebElement> alllist = driver.FindElements(By.XPath(ele));
            IList<IWebElement> allsublist = new List<IWebElement>(alllist);
            for (int i = 0; i < alllist.Count; i++)
            {
                //string startupPath = Directory.GetCurrentDirectory().Remove("\\redbus\\Output\\");
                 //startupPath = null;
                string workingDirectory = Environment.CurrentDirectory;
                if (workingDirectory != null)
                {
                    String startupPath = Directory.GetParent(workingDirectory).Parent.Parent.FullName;
                    if (!Directory.Exists(startupPath+"\\Output"))
                    {
                        Directory.CreateDirectory((startupPath + "\\Output"));
                    }
                    String paths = startupPath + "\\output\\" + path + "_operator.txt";
                    DirectoryInfo di = Directory.CreateDirectory(path);
                    Console.WriteLine("The directory was created successfully");

                    //using (FileStream fa = new FileStream(@paths, FileMode.Truncate, FileAccess.Write))
                    //{
                    //    //fa.Flush();
                    //    fa.Close();
                    //}
                    FileStream fs = new FileStream(@paths, FileMode.Append, FileAccess.Write, FileShare.Write);

                    fs.Close();

                    StreamWriter sw = new StreamWriter(@paths, true, Encoding.ASCII);
                    //Console.WriteLine(alllist[i].Text);
                    sw.Write(string.Empty);
                    JSExecutorHelper.jsjs(driver, alllist[i]);
                    sw.WriteLine(alllist[i].Text);
                    sw.Close();
                    //File.Create(paths).Close();
                }
            }
        }
        public static void switchtabs(IWebDriver driver)
        {
            Thread.Sleep(1000);
            var tabs = driver.WindowHandles;
            if (tabs.Count > 1)
            {
                driver.SwitchTo().Window(tabs[1]);
            }
        }

        public static String CurrentTestName()
        {
            string name = TestContext.CurrentContext.Test.Name;
            Console.WriteLine("Name of the method - " + name);
            return name;
        }


        public static void createDirectory(String pathToCreate,string fextension)
        {
            string workingDirectory = Environment.CurrentDirectory;
            if (workingDirectory != null)
            {
                String startupPath = Directory.GetParent(workingDirectory).Parent.Parent.FullName;
                if (!Directory.Exists(startupPath + "\\"+ pathToCreate))
                {
                    Directory.CreateDirectory((startupPath + "\\" + pathToCreate));
                }
                String paths = startupPath + "\\" + pathToCreate + "\\"+fextension;
                DirectoryInfo di = Directory.CreateDirectory(paths);
                Console.WriteLine("The directory was created successfully");

            }
        }
}
