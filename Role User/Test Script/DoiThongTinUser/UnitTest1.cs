using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;

namespace DoiThongTinUser
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\DoiThongTinUser\ExcelData\Data.xlsx";
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string user = worksheet.Cells[row, 1].Value.ToString();
                    string pass = worksheet.Cells[row, 2].Value.ToString();
                    string name = worksheet.Cells[row,3].Value.ToString();
                    string phonenum = worksheet.Cells[row, 4].Value.ToString();
                    string gioitinh = worksheet.Cells[row, 5].Value.ToString();
                    string diachi = worksheet.Cells[row, 6].Value.ToString();

                    if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(phonenum) || string.IsNullOrEmpty(gioitinh) || string.IsNullOrEmpty(diachi))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue;
                    }

                    DoiThongTin(user, pass, name,phonenum,gioitinh,diachi);

                    bool isErrorMessageDisplayed = false;

                    try
                    {
                        IWebElement ketqua = driver.FindElement(By.XPath("//button[@class='search__submit button--link']"));
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\DoiThongTinUser\ExcelData\Data.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string username = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string tenuser = worksheet.Cells[row, 3].Value.ToString();
                    string phone = worksheet.Cells[row, 4].Value.ToString();
                    string gioitinh = worksheet.Cells[row, 5].Value.ToString();
                    string diachi = worksheet.Cells[row, 6].Value.ToString();

                    yield return new string[] { username, password,tenuser,phone,gioitinh,diachi};
                }
            }
        }
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 7].Value = result;

        }
        public void DoiThongTin(string username, string password, string tenuser,string sodth, string gioitinh,string diachi)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/");
            driver.FindElement(By.Id("btn-login")).Click();
            Thread.Sleep(1000);
            IWebElement taikhoan = driver.FindElement(By.CssSelector("#login_email"));
            taikhoan.SendKeys(username);
            Thread.Sleep(1000);
            IWebElement matkhau = driver.FindElement(By.CssSelector("#login_password"));
            matkhau.SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//button[@id='form_btn-login']")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".user__info--name > span")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Thông tin cá nhân")).Click();
            Thread.Sleep(1000);
            IWebElement TenUser = driver.FindElement(By.CssSelector("#TenUngVien"));
            TenUser.Clear();
            TenUser.Click();
            TenUser.SendKeys(tenuser);
            Thread.Sleep(1000);
            IWebElement SoDT = driver.FindElement(By.CssSelector("#SoDienThoai"));
            SoDT.Clear();
            SoDT.Click();
            SoDT.SendKeys(sodth);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("GioiTinh")).Click();
            {
                var dropdown = driver.FindElement(By.Id("GioiTinh"));
                dropdown.FindElement(By.XPath("//option[. = '"+gioitinh+"']")).Click();
            }Thread.Sleep(1000);
            Thread.Sleep(1000);
            IWebElement address = driver.FindElement(By.CssSelector("#DiaChi"));
            address.Clear();
            address.Click();
            address.SendKeys(diachi);
            Thread.Sleep(1000);
            IWebElement updatebtn = driver.FindElement(By.XPath("//input[@value='Cập nhật thông tin']"));
            updatebtn.Click();  
        }
    }
}