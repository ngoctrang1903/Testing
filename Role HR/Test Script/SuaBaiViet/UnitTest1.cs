using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;

namespace SuaBaiViet
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\SuaBaiViet\ExcelData\SuaBaiViet.xlsx";
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
                    string user = worksheet.Cells[row, 1].Value.ToString();
                    string pass = worksheet.Cells[row, 2].Value.ToString();
                    string name = worksheet.Cells[row, 3].Value.ToString();
                    string content =worksheet.Cells[row, 4].Value.ToString();
                    if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(content))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue;
                    }
                    EditBV(user, pass, name, content);

                    bool isErrorMessageDisplayed = false;

                    try
                    {
                        IWebElement ketqua = driver.FindElement(By.XPath("//tbody/tr[1]/td[7]/span[1]"));
                        ketqua.Click();

                        isErrorMessageDisplayed = true;
                        if (ketqua == null)
                        {
                            isErrorMessageDisplayed = true;
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\SuaBaiViet\ExcelData\SuaBaiViet.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string username = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string tenbv = worksheet.Cells[row, 3].Value.ToString();
                    string noidung = worksheet.Cells[row, 4].Value.ToString();

                    yield return new string[] { username, password, tenbv, noidung };
                }
            }
        }
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 5].Value = result;

        }
        public void EditBV(string username, string password, string tenbv, string noidung)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/nha-tuyen-dung");

            IWebElement usernameInput = driver.FindElement(By.CssSelector("#Email"));
            usernameInput.SendKeys(username);
            Thread.Sleep(2000);
            IWebElement passwordInput = driver.FindElement(By.CssSelector("#Password"));
            passwordInput.SendKeys(password);
            Thread.Sleep(2000);

            IWebElement loginbtn = driver.FindElement(By.XPath("//button[contains(text(),'Đăng nhập')]"));
            loginbtn.Click();
            Thread.Sleep(2000);

            driver.FindElement(By.CssSelector("body > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(5) > a:nth-child(1)")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//span[contains(text(),'Bài viết chờ duyệt')]")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/BaiViet/Edit/8']")).Click();
            Thread.Sleep(2000);
            IWebElement TenBV = driver.FindElement(By.CssSelector("#TenBaiViet"));
            TenBV.SendKeys(tenbv);
            Thread.Sleep(2000);
            driver.SwitchTo().Frame(0);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + noidung + "'}", element);
            }
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//input[@value='Cập nhật']")).Click();
            Thread.Sleep(2000);
        }

    }
}