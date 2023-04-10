using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject_framework.Helpers
{
    public class BrowserHelpers:DriverHelper
    {
        public IWebDriver ChromeInitialize(IWebDriver driver, String url)
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(url);
            return driver;
        }
    }
}
