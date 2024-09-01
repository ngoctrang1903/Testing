using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;
using OpenQA.Selenium.Edge;
using System.Xml.Linq;

namespace RegisterUser
{
    public class Tests
    {
        public IWebDriver driver;

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
        public void EditWithExcelData()
        {
            string filePath = @"C:\Users\Hoang Phuc\Desktop\RegisterUser\ExcelData\TestInput.xlsx";
            FileInfo file = new FileInfo(filePath);
            if (!file.Exists)
            {
                Console.WriteLine("Excel file does not exist.");
                return;
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                {
                    Console.WriteLine("No worksheet found in the Excel file.");
                    return;
                }
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string nameInput = worksheet.Cells[row, 1].Value.ToString();
                    string emailInput = worksheet.Cells[row, 2].Value.ToString();
                    string passInput = worksheet.Cells[row, 3].Value.ToString();
                    string passConfirmInput = worksheet.Cells[row, 4].Value.ToString();

                    if (string.IsNullOrEmpty(nameInput) || string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passInput) || string.IsNullOrEmpty(passConfirmInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue;
                    }
                    RegisterUser(nameInput, emailInput, passInput, passConfirmInput);

                    bool error = true;
                    IWebElement ketqua = driver.FindElement(By.CssSelector("h4[class='note-title']"));
                    try
                    {
                        if (ketqua==null)
                        {
                            error = true;
                        }
                        else
                        {
                            error = false;
                        }
                    }
                    catch (NoSuchDriverException)
                    {
                        error = false;
                    }


                    string result = error ? "Pass" : "Fail";

                    UpdateExcelResult(worksheet, row, result);

                }

                package.Save();
            }
        }
        private static IEnumerable<string[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"C:\Users\Hoang Phuc\Desktop\RegisterUser\ExcelData\TestInput.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string name = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 2].Value.ToString();
                    string password = worksheet.Cells[row, 3].Value.ToString();
                    string passwordconfirm = worksheet.Cells[row, 4].Value.ToString();

                    yield return new string[] { name, email, password, passwordconfirm };
                }
            }
        }
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 5].Value = result;

        }
        public void RegisterUser(string name, string email, string password, string passwordconfirm)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/");
            driver.FindElement(By.CssSelector("#btn-register")).Click();
            Thread.Sleep(2000);

            driver.FindElement(By.CssSelector("#register_name")).SendKeys(name);
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("#register_email")).SendKeys(email);
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("#register_password")).SendKeys(password);
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("#password_confirm")).SendKeys(passwordconfirm);
            Thread.Sleep(2000);
            IWebElement registerbtn = driver.FindElement(By.CssSelector("#form_btn-register"));
            registerbtn.Click();

        }

    }
}