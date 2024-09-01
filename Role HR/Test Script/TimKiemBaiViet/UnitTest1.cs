using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;

namespace TimKiemBaiViet
{
    [TestFixture]
    public class SearchBaiViet
    {
        private IWebDriver driver;
        private string baseUrl = "http://localhost:62536/nha-tuyen-dung";

        [SetUp]
        public void SetupTest()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
        }

        [TearDown]
        public void TeardownTest()
        {
            driver.Quit();
        }

        [Test]
        public void RunTestCasesFromExcel()
        {
            string inputFilePath = @"C:\Users\Hoang Phuc\Desktop\TimKiemBaiViet\ExcelData\TestInput.xlsx";
            string outputFilePath = @"C:\Users\Hoang Phuc\Desktop\TimKiemBaiViet\ExcelData\TestOutput.xlsx";


            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file '{inputFilePath}' not found.");
                return;
            }


            FileInfo inputFile = new FileInfo(inputFilePath);
            using (ExcelPackage package = new ExcelPackage(inputFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;


                FileInfo outputFile = new FileInfo(outputFilePath);
                using (ExcelPackage outputPackage = new ExcelPackage(outputFile))
                {

                    string resultSheetName = $"Test Results_{DateTime.Now:yyyy:MM:dd:HH:mm}";
                    ExcelWorksheet outputWorksheet = outputPackage.Workbook.Worksheets.Add(resultSheetName);


                    outputWorksheet.Cells[1, 1].Value = "Email";
                    outputWorksheet.Cells[1, 2].Value = "Password";
                    outputWorksheet.Cells[1, 3].Value = "Keyword";
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string email = worksheet.Cells[row, 1].Value?.ToString();
                        string password = worksheet.Cells[row, 2].Value?.ToString();
                        string keyword = worksheet.Cells[row, 3].Value?.ToString();


                        if (!string.IsNullOrEmpty(email) && IsValidEmail(email))
                        {

                            string result = SearchTest(email, password, keyword);

                            outputWorksheet.Cells[row, 1].Value = email;
                            outputWorksheet.Cells[row, 2].Value = password;
                            outputWorksheet.Cells[row, 3].Value = keyword;

                            outputWorksheet.Cells[row, 4].Value = result;
                        }
                        else
                        {
                            outputWorksheet.Cells[row, 1].Value = email;
                            outputWorksheet.Cells[row, 2].Value = keyword;
                            outputWorksheet.Cells[row, 3].Value = keyword;
                            outputWorksheet.Cells[row, 4].Value = "Fail";
                        }
                    }
                    outputPackage.Save();
                }
            }
        }

        public bool IsValidEmail(string email)
        {
            try
            {
                MailAddress mailAddress = new MailAddress(email);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        public string SearchTest(string email, string password, string keyword)
        {
            driver.Navigate().GoToUrl(baseUrl);
            driver.FindElement(By.CssSelector("#Email")).SendKeys(email);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("#Password")).SendKeys(password);
            Thread.Sleep(1000);

            IWebElement loginbtn = driver.FindElement(By.CssSelector("button[type='submit']"));
            loginbtn.Click();
            Thread.Sleep(2000);

            IWebElement baiviet = driver.FindElement(By.CssSelector("body > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(5) > a:nth-child(1)"));
            baiviet.Click();
            Thread.Sleep(2000);
            IWebElement baivietcuaban = driver.FindElement(By.CssSelector("li[class='mm-active'] li:nth-child(1) a:nth-child(1)"));
            baivietcuaban.Click();
            Thread.Sleep(2000);

            IWebElement textSearch = driver.FindElement(By.XPath("//input[@id='txtsearch']"));
            textSearch.SendKeys(keyword);

            IWebElement ketquatimkiem = driver.FindElement(By.Id("selection-datatable_info"));
            if (ketquatimkiem.Text == "1/0")
            {
                return "Fail";
            }
            else
            {
                return "Pass";
            }
        }
    }
}