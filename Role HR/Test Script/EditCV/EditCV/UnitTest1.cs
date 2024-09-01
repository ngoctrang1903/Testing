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
namespace EditCV

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
            string excelFilePath = @"C:\Users\Admin\Desktop\EditCV\editCV.xlsx";
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
                    string nameCVInput = worksheet.Cells[row, 3].Value?.ToString();
                    string baccvInput = worksheet.Cells[row, 4].Value?.ToString();
                    string chuyennganhInput = worksheet.Cells[row, 5].Value?.ToString();
                    string loaiCVInput = worksheet.Cells[row, 6].Value?.ToString();
                    string muctieuInput = worksheet.Cells[row, 7].Value?.ToString();
                    string kinhnghiemInput = worksheet.Cells[row, 8].Value?.ToString();
                    string kinangInput = worksheet.Cells[row, 9].Value?.ToString();
                    string hocvanInput = worksheet.Cells[row, 10].Value?.ToString();
                    string kinangmemInput = worksheet.Cells[row, 11].Value?.ToString();
                    string giaithuongInput = worksheet.Cells[row, 12].Value?.ToString();



                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(nameCVInput) || string.IsNullOrEmpty(baccvInput) || string.IsNullOrEmpty(chuyennganhInput) || string.IsNullOrEmpty(loaiCVInput) || string.IsNullOrEmpty(muctieuInput) || string.IsNullOrEmpty(kinhnghiemInput) || string.IsNullOrEmpty(kinangInput) || string.IsNullOrEmpty(hocvanInput) || string.IsNullOrEmpty(kinangmemInput) || string.IsNullOrEmpty(giaithuongInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel
                    EditCV(emailInput, passwordInput, nameCVInput, baccvInput, chuyennganhInput, loaiCVInput, muctieuInput, kinhnghiemInput, kinangInput, hocvanInput, kinangmemInput, giaithuongInput);

                    bool isErrorMessageDisplayed = true;
                    try
                    {
                        IWebElement errorMessage1 = driver.FindElement(By.XPath("//li[contains(text(),'Tên hồ sơ đã tồn tại')]"));

                        if (errorMessage1.Text == "Tên hồ sơ đã tồn tại")
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
            string filePath = @"C:\Users\Admin\Desktop\EditCV\editCV.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string tenhs = worksheet.Cells[row, 3].Value.ToString();
                    string baccv = worksheet.Cells[row, 4].Value.ToString();
                    string chuyennganh = worksheet.Cells[row, 5].Value.ToString();
                    string loaicv = worksheet.Cells[row, 6].Value.ToString();
                    string muctieu = worksheet.Cells[row, 7].Value.ToString();
                    string exp = worksheet.Cells[row, 8].Value.ToString();
                    string kinang = worksheet.Cells[row, 9].Value.ToString();
                    string hocvan = worksheet.Cells[row, 10].Value.ToString();
                    string kinangmem = worksheet.Cells[row, 11].Value.ToString();
                    string giaithuong = worksheet.Cells[row, 12].Value.ToString();

                    yield return new string[] { email, password, tenhs, baccv, chuyennganh, loaicv, muctieu, exp, kinang, hocvan, kinangmem, giaithuong };
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 13].Value = result;
        }
        private void EditCV(string email, string password, string tenhs, string baccv, string chuyennganh, string loaicv, string muctieu, string exp, string kinang, string hocvan, string kinangmem, string giaithuong)
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
            Thread.Sleep(2000);
            driver.FindElement(By.LinkText("Quản lý CV")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.LinkText("Sửa")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenHoSo")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenHoSo")).SendKeys(tenhs);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".col-md-4:nth-child(1) .filter-option-inner-inner")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("li:nth-child(2) .text")).Click();
            Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("MaCapBac"));
                dropdown.FindElement(By.XPath("//option[. = '"+ baccv +"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".col-md-4:nth-child(2) .filter-option-inner-inner")).Click();
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".col-md-4:nth-child(2) .filter-option-inner-inner"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element).Perform();
            }
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.TagName("body"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element, 0, 0).Perform();
            }
            Thread.Sleep(1000);
            //driver.FindElement(By.LinkText("C/C++")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaCN"));
                dropdown.FindElement(By.XPath("//option[. = '"+ chuyennganh+"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".col-md-4:nth-child(3) .filter-option-inner-inner")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".col-md-4:nth-child(3) li:nth-child(4) .text")).Click();
            Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("MaLoaiCV"));
                dropdown.FindElement(By.XPath("//option[. = '"+loaicv+"']")).Click();
            }
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(0);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+muctieu+"'}", element);
            }
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+exp+"'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(1);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+kinang+"'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(2);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+hocvan+"'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(3);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();
            Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+kinangmem+"'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(4);
            Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector("html")).Click();
            //Thread.Sleep(1000);
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '"+kinangmem+"'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("input[value='Cập nhật hồ sơ']")).Click();
            Thread.Sleep(1000);
        }
    }
}