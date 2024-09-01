using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace CreateUV
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
        public void CreateWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\CreateUV\createUV.xlsx";
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
                    string idInput = worksheet.Cells[row, 3].Value?.ToString();
                    string nameInput = worksheet.Cells[row, 4].Value?.ToString();
                    string phoneInput = worksheet.Cells[row, 5].Value?.ToString();
                    string sexInput = worksheet.Cells[row, 6].Value?.ToString();
                    string addressInput = worksheet.Cells[row, 7].Value?.ToString();


                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(idInput) || string.IsNullOrEmpty(nameInput)
                        || string.IsNullOrEmpty(phoneInput) || string.IsNullOrEmpty(sexInput)
                        || string.IsNullOrEmpty(addressInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel
                    Create(emailInput, passwordInput, idInput, nameInput, phoneInput, sexInput, addressInput);
                    bool isErrorMessageDisplayed = true;
                    try
                    {
                        IWebElement errorMessage1 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors text-danger'] ul li"));
                        if (errorMessage1.Text == "Mã ứng viên đã tồn tại" )
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    try
                    {
                        IWebElement errorMessage2 = driver.FindElement(By.CssSelector(".text-danger.field-validation-error"));
                        if (errorMessage2.Text == "The field MaUngVien must be a number.")
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
            string filePath = @"C:\Users\Admin\Desktop\CreateUV\createUV.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value?.ToString();
                    string password = worksheet.Cells[row, 2].Value?.ToString();
                    string id = worksheet.Cells[row, 3].Value?.ToString();
                    string name = worksheet.Cells[row, 4].Value?.ToString();
                    string phone = worksheet.Cells[row, 5].Value?.ToString();
                    string sex = worksheet.Cells[row, 6].Value?.ToString();
                    string address = worksheet.Cells[row, 7].Value?.ToString();

                    yield return new string[] { email, password, id, name, phone, sex,  address };
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 8].Value = result;

        }
        private void Create(string email, string password, string id, string name, string phone, string sex,  string address)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/Admin/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1296, 696);
            driver.FindElement(By.CssSelector(".card-body")).Click();
            driver.FindElement(By.Id("UserName")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("UserName")).SendKeys(email);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("PassWord")).Click();
            driver.FindElement(By.Id("PassWord")).SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//button[contains(text(),'Đăng nhập')]")).Click();
            Thread.Sleep(1000);
            js.ExecuteScript("window.scrollTo(0,0)");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("li:nth-child(4) span:nth-child(2)")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Thêm ứng viên")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MaUngVien")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MaUngVien")).SendKeys(id);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenUngVien")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenUngVien")).SendKeys(name);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("SoDienThoai")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("SoDienThoai")).SendKeys(phone);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("GioiTinh")).Click();
            Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("GioiTinh"));
                dropdown.FindElement(By.XPath("//option[. = '"+ sex +"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.Id("DiaChi")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("DiaChi")).SendKeys(address);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector(".col-xl-12:nth-child(1)")).Click();
        }
    }
}
