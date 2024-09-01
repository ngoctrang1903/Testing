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

namespace AddNTD
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
        public void AddNTDWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\AddNTD\addNTD.xlsx";
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
                    string maNTDInput = worksheet.Cells[row, 3].Value?.ToString();
                    string tenNTDInput = worksheet.Cells[row, 4].Value?.ToString();
                    string tenNDDInput = worksheet.Cells[row, 5].Value?.ToString();
                    string vitriInput = worksheet.Cells[row, 6].Value?.ToString();
                    string phoneInput = worksheet.Cells[row, 7].Value?.ToString();
                    string quymoInput = worksheet.Cells[row, 8].Value?.ToString();
                    string motaInput = worksheet.Cells[row, 9].Value?.ToString();
                    string diachiInput = worksheet.Cells[row, 10].Value?.ToString();
                    string tenwebInput = worksheet.Cells[row, 11].Value?.ToString();

                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(maNTDInput) || string.IsNullOrEmpty(tenNTDInput) 
                        || string.IsNullOrEmpty(tenNDDInput) || string.IsNullOrEmpty(vitriInput) || string.IsNullOrEmpty(phoneInput) || string.IsNullOrEmpty(quymoInput) 
                        || string.IsNullOrEmpty(motaInput) || string.IsNullOrEmpty(diachiInput) || string.IsNullOrEmpty(tenwebInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel

                    AddNTD(emailInput, passwordInput, maNTDInput, tenNTDInput, tenNDDInput, vitriInput, phoneInput, quymoInput, motaInput, diachiInput, tenwebInput);

                    bool isErrorMessageDisplayed = true;
                    try
                    {
                        IWebElement errorMessage1 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors text-danger'] ul li"));

                        if (errorMessage1.Text == "Mã nhà tuyển dụng đã tồn tại")
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    try
                    {
                        IWebElement errorMessage2 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors text-danger'] ul li"));

                        if (errorMessage2.Text == "Mã nhà tuyển chưa trùng với mã tài khoản đăng ký nhà tuyển dụng")
                        {
                            isErrorMessageDisplayed = false;
                        }

                    }
                    catch (NoSuchElementException)
                    {

                    }

                    try
                    {
                        IWebElement errorMessage3 = driver.FindElement(By.CssSelector("#MaNTD-error"));

                        if (errorMessage3.Text == "The field MaNTD must be a number.")
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
            string filePath = @"C:\Users\Admin\Desktop\AddNTD\addNTD.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value?.ToString();
                    string password = worksheet.Cells[row, 2].Value?.ToString();
                    string maNTD = worksheet.Cells[row, 3].Value?.ToString();
                    string tenNTD = worksheet.Cells[row, 4].Value?.ToString();
                    string tenNDD = worksheet.Cells[row, 5].Value?.ToString();
                    string vitri = worksheet.Cells[row, 6].Value?.ToString();
                    string phone = worksheet.Cells[row, 7].Value?.ToString();
                    string quymo = worksheet.Cells[row, 8].Value?.ToString();
                    string mota = worksheet.Cells[row, 9].Value?.ToString();
                    string diachi = worksheet.Cells[row, 10].Value?.ToString();
                    string tenweb = worksheet.Cells[row, 11].Value?.ToString();

                    yield return new string[] { email, password, maNTD, tenNTD, tenNDD, vitri, phone, quymo, mota, diachi, tenweb};
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 12].Value = result;

        }
        private void AddNTD(string email, string password, string maNTD, string tenNTD, string tenNDD, string vitri, string phone, string quymo, string mota, string diachi, string tenweb)
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
            driver.FindElement(By.CssSelector("li:nth-child(5) span:nth-child(2)")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Thêm nhà tuyển dụng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MaNTD")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MaNTD")).SendKeys(maNTD);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenNTD")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenNTD")).SendKeys(tenNTD);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenNDD")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TenNDD")).SendKeys(tenNDD);
            Thread.Sleep(1000);
            //driver.FindElement(By.Id("ChucVuNDD")).Click();
            //Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("ChucVuNDD"));
                dropdown.FindElement(By.XPath("//option[. = '"+ vitri+"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.Id("SoDienThoai")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("SoDienThoai")).SendKeys(phone);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("QuyMo")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("QuyMo")).SendKeys(quymo);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MoTa")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("MoTa")).SendKeys(mota);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("DiaChi")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("DiaChi")).SendKeys(diachi);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("Website")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("Website")).SendKeys(tenweb);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
        }
    }
}
