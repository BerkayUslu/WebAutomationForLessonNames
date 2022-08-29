using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;



namespace GetLessonNames
{

    class MainClass
    {
        static List<string> lessonData = new List<string>();
        static string lessonName;
        static string programName;

        public static void Main(string[] args)
        {
            ChromeDriver driver = Initiate();
            ScanSites(driver);
            WriteToExcell();

            driver.Quit();
        }

        private static void WriteToExcell()
        {
            float i = 0;
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            foreach (string data in lessonData)
            {
                var index = Math.Truncate(i / 3) + 1;
                if (i % 3 == 0) { worksheet.Cell("A" + index).Value = data; }
                else if (i % 3 == 1) { worksheet.Cell("B" + index).Value = data; }
                else { worksheet.Cell("C" + index).Value = data; }
                i++;
            }

            workbook.SaveAs("HelloWorld.xlsx");
        }

        private static void ScanSites(ChromeDriver driver)
        {
            for (int i = 15189; i <= 15210; i++)
            {
                driver.Navigate().GoToUrl("https://online.yildiz.edu.tr/?transaction=LMS.EDU.LessonProgram.ViewOnlineLessonProgramForStudent/" + i);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
                var nameElements = driver.FindElements(By.ClassName("col-sm-8"));
                if (nameElements.Count == 0) { continue; }
                lessonName = nameElements[0].Text;
                programName = nameElements[1].Text;
                lessonData.Add(i.ToString());
                lessonData.Add(lessonName);
                lessonData.Add(programName);
            }
        }

        private static ChromeDriver Initiate()
        {
            var options = new ChromeOptions();
            var driver = new ChromeDriver(options);

            driver.Navigate().GoToUrl("https://online.yildiz.edu.tr/Account/Login?ReturnUrl=%2f");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);

            var mail = driver.FindElement(By.Name("Data.Mail"));
            var password = driver.FindElement(By.Name("Data.Password"));
            var submitButton = driver.FindElement(By.ClassName("btn"));

            //Enter mail and password
            mail.SendKeys("");
            password.SendKeys("");
            submitButton.Click();

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
            return driver;
        }
    }
}
