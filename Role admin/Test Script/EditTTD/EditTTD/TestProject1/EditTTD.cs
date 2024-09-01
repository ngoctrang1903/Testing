using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V120.Input;
using OpenQA.Selenium.DevTools.V120.Network;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace EditTTD
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
        public void EditTTDWithExcelData()
        {
            // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
            string excelFilePath = @"C:\Users\Admin\Desktop\EditTTD\editTTD.xlsx";
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
                    string capbacInput = worksheet.Cells[row, 3].Value?.ToString();
                    string hinthucInput = worksheet.Cells[row, 4].Value?.ToString();
                    string trangthaiInput = worksheet.Cells[row, 5].Value?.ToString();
                    string noidungInput = worksheet.Cells[row, 6].Value?.ToString();


                    if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(capbacInput) || string.IsNullOrEmpty(hinthucInput) 
                        || string.IsNullOrEmpty(trangthaiInput) || string.IsNullOrEmpty(noidungInput))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue; // Skip to the next row
                    }
                    // Test registration with the data from Excel

                    EditTTD(emailInput, passwordInput, capbacInput, hinthucInput, trangthaiInput, noidungInput);

                    bool isErrorMessageDisplayed = false;
                    try
                    {
                        IWebElement ketqua = driver.FindElement(By.CssSelector("tbody tr:nth-child(2) td:nth-child(9) a:nth-child(1)"));
                        
                        isErrorMessageDisplayed = true;
                        if (ketqua != null)
                        {

                            isErrorMessageDisplayed = true;
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
            string filePath = @"C:\Users\Admin\Desktop\EditTTD\editTTD.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string capbac = worksheet.Cells[row, 3].Value.ToString();
                    string hinhthuc = worksheet.Cells[row, 4].Value?.ToString();
                    string trangthai = worksheet.Cells[row, 5].Value?.ToString();
                    string noidung = worksheet.Cells[row, 6].Value?.ToString();

                    yield return new string[] { email, password, capbac, hinhthuc, trangthai, noidung };
                }
            }
        }
        // Update the Excel file with the result
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 7].Value = result;

        }
        private void EditTTD(string email, string password, string capbac, string hinhthuc, string trangthai, string noidung)
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
            driver.FindElement(By.CssSelector("li:nth-child(5) > .waves-effect > span:nth-child(2)")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Danh sách tin tuyển dụng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Cập nhật")).Click();
            Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector(".form-group:nth-child(3) > .col-md-4:nth-child(2) .filter-option-inner-inner")).Click();
            //Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn cấp bậc --']")).Click();
            //Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("MaCapBac"));
                dropdown.FindElement(By.XPath("//option[. = '" + capbac +"']")).Click();
            }
            Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector(".form-group:nth-child(4) > .col-md-4:nth-child(2) .filter-option-inner-inner")).Click();
            //Thread.Sleep(1000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn hình thức làm việc --']")).Click();
            //Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("MaLoaiCV"));
                dropdown.FindElement(By.XPath("//option[. = '"+ hinhthuc +"']")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.Id("SoLuong")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("TrangThai")).Click();
            Thread.Sleep(1000);
            {
                var dropdown = driver.FindElement(By.Id("TrangThai"));
                dropdown.FindElement(By.XPath("//option[. = '"+ trangthai+"']")).Click();
            }
            Thread.Sleep(1000);
            driver.SwitchTo().Frame(4);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("html")).Click();

            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + noidung + "'}", element);
            }
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("input[value='Cập nhật']")).Click();
            Thread.Sleep(1000);
        }
    }
}
