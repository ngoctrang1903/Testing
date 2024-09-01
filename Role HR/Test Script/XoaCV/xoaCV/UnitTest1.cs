using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;
using OpenQA.Selenium.Interactions;
using System.Timers;
namespace xoaCV

{
    public class Tests
    {
        private IWebDriver driver;

        public IDictionary<string, object> vars { get; private set; }
        private IJavaScriptExecutor js;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<string, object>();

            driver.Manage().Window.Maximize();
        }

        [TearDown]
        public void Teardown()
        {
            driver.Quit();
        }

        [Test]
        public void EditCVWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\XoaCV\xoaCV.xlsx";
            // Load Excel file
            FileInfo fileInfo = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming your data is in the first sheet
                

                // Start from the second row (assuming the first row is headers)
                int rowCount = worksheet.Dimension.Rows;

                // Start from the second row (assuming the first row is headers)
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    string emailInput = worksheet.Cells[row, 1].Value?.ToString();
                    string passwordInput = worksheet.Cells[row, 2].Value?.ToString();
                   



                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel
                    EditCV(emailInput, passwordInput);

                    bool isErrorMessageDisplayed = true;
                    try
                    {
                        IWebElement errorMessage1 = driver.FindElement(By.XPath("//a[contains(text(),'Xem hồ sơ')]"));

                        if (errorMessage1 == null)
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    string result = isErrorMessageDisplayed ? "Pass" : "Fail";

                    UpdateExcelResult(worksheet, row, result);
                }
                package.Save();
            }

        }
        private static IEnumerable<string[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"C:\Users\Admin\Desktop\XoaCV\xoaCV.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                   

                    yield return new string[] { email, password };
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 3].Value = result;
        }
        private void EditCV(string email, string password)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/");
            driver.Manage().Window.Size = new System.Drawing.Size(1296, 696);
            if (driver.FindElements(By.CssSelector(".user__info--name > span")).Count == 0)
            {
                // Nếu chưa đăng nhập, thực hiện các bước đăng nhập
                driver.FindElement(By.Id("btn-login")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".main__login--content")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("login_email")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("login_email")).SendKeys(email);
                Thread.Sleep(1000);
                driver.FindElement(By.Id("login_password")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("login_password")).SendKeys(password);
                Thread.Sleep(1000);
                driver.FindElement(By.Id("form_btn-login")).Click();
                Thread.Sleep(1000);
            }
            driver.FindElement(By.LinkText("Hồ sơ & CV")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Quản lý CV")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//a[normalize-space()='Xóa']"));
        }
    }
}