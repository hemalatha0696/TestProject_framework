using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenQA.Selenium.Support.UI;
namespace TestProject_framework.Helpers
{
    public class JSExecutorHelper:DriverHelper
    {
        public static void totop(IWebDriver driver)
        {
            IJavaScriptExecutor jsjsall = (IJavaScriptExecutor)driver;
            jsjsall.ExecuteScript("window.scrollTo(0, 0);");
        }
        public static void js(IWebDriver driver)
        {
            IJavaScriptExecutor jsjsall = (IJavaScriptExecutor)driver;
            jsjsall.ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
        }
        public static void jsjs(IWebDriver driver, IWebElement ae)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", ae);
        }
        public static void scroll(IWebDriver driver, String coordinates)
        {

            IJavaScriptExecutor scrd = (IJavaScriptExecutor)(driver);
            scrd.ExecuteScript(coordinates);

        }
    }
}
