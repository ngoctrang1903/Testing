using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace SearchTK
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
        public void EditBVWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\EditBV\EditBV.xlsx";
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
                    string trangthaiInput = worksheet.Cells[row, 3].Value?.ToString();
                   
                  
                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(trangthaiInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel
                    EditBV(emailInput, passwordInput, trangthaiInput);

                    bool isErrorMessageDisplayed = false;

                    try
                    {
                        IWebElement ketqua = driver.FindElement(By.CssSelector("tbody tr:nth-child(1) td:nth-child(8) a:nth-child(1)"));
                        ketqua.Click();
                        isErrorMessageDisplayed = true;
                        if (ketqua == null)
                        {
                            isErrorMessageDisplayed= true;
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
            string filePath = @"C:\Users\Admin\Desktop\EditBV\EditBV.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string trangthai = worksheet.Cells[row, 3].Value.ToString();
                 
                    yield return new string[] { email, password, trangthai};
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 4].Value = result;

        }
        private void EditBV(string email, string password, string trangthai)
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
            driver.FindElement(By.CssSelector("li:nth-child(7) > .waves-effect")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Danh sách bài viết")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Cập nhật")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TrangThai")).Click();
            Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("TrangThai"));
                dropdown.FindElement(By.XPath("//option[. = '"+ trangthai +"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("input[value='Cập nhật']")).Click();
            Thread.Sleep(1000);
        }
    }
}
