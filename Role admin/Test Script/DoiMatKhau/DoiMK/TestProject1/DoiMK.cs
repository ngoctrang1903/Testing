using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V120.Input;
using OpenQA.Selenium.DevTools.V120.Network;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace DoiMK
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
        public void DoiMKWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\DoiMatKhau\doiMK.xlsx";
            // Load Excel file
            FileInfo fileInfo = new FileInfo(excelFilePath);
            if (!fileInfo.Exists)
            {
                Console.WriteLine("Excel file does not exist.");
                return;
            }
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming your data is in the first sheet
                if (worksheet == null)
                {
                    Console.WriteLine("No worksheet found in the Excel file.");
                    return;
                }

                // Start from the second row (assuming the first row is headers)
                int rowCount = worksheet.Dimension.Rows;

                // Start from the second row (assuming the first row is headers)
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    string emailInput = worksheet.Cells[row, 1].Value?.ToString();
                    string passwordInput = worksheet.Cells[row, 2].Value?.ToString();
                    string oldpassInput = worksheet.Cells[row, 3].Value?.ToString();
                    string newpassInput = worksheet.Cells[row, 4].Value?.ToString();
                    string confirmpassInput = worksheet.Cells[row, 5].Value?.ToString();


                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(oldpassInput) || string.IsNullOrEmpty(newpassInput)
                        || string.IsNullOrEmpty(confirmpassInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel

                    DoiMK(emailInput, passwordInput, oldpassInput, newpassInput, confirmpassInput);

                    bool isErrorMessageDisplayed = true;
                    try
                    {
                        IWebElement errorMessage1 = driver.FindElement(By.XPath("//li[contains(text(),'Mật khẩu cũ chưa chính xác')]"));

                        if (errorMessage1.Text == "Mật khẩu cũ chưa chính xác")
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    try
                    {
                        IWebElement errorMessage2 = driver.FindElement(By.XPath("//span[@class='field-validation-error text-danger']"));

                        if (errorMessage2.Text == "Mật khẩu xác nhận chưa đúng")
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    string result = isErrorMessageDisplayed ? "Pass" : "Fail";

                    UpdateExcelResult(worksheet,row, result);


                }
                package.Save();
            }

        }
        private static IEnumerable<string[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"C:\Users\Admin\Desktop\DoiMatKhau\doiMK.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value?.ToString();
                    string password = worksheet.Cells[row, 2].Value?.ToString();
                    string oldpass = worksheet.Cells[row, 3].Value?.ToString();
                    string newpass = worksheet.Cells[row, 4].Value?.ToString();
                    string confirmpass = worksheet.Cells[row, 5].Value?.ToString();
                    

                    yield return new string[] { email, password, oldpass, newpass, confirmpass};
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 6].Value = result;

        }
        private void DoiMK(string email, string password, string oldpass, string newpass, string confirmpass)
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
            driver.FindElement(By.CssSelector(".user__info--name > span")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Đổi mật khẩu")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MatKhauCu")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MatKhauCu")).SendKeys(oldpass);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MatKhauMoi")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MatKhauMoi")).SendKeys(newpass);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("XN_MatKhau")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("XN_MatKhau")).SendKeys(confirmpass);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("input[value='Cập nhật']")).Click();
            Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector(".user__info--name > span")).Click();
            //Thread.Sleep(1000);
            //driver.FindElement(By.LinkText("Đăng xuất")).Click();
            //Thread.Sleep(2000);
        }
    }
}
