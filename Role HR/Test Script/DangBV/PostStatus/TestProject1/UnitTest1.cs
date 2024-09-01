using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Numerics;
using System.Xml.Linq;

[TestFixture]
public class ExcelDataDrivenTest
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
    }

    [TearDown]
    public void Teardown()
    {
        driver.Quit();
    }

    [Test]
    public void PostStatusWithExcelData()
    {
        // Đường dẫn đến tệp Excel chứa dữ liệu đăng bài viết
        string excelFilePath = @"C:\Users\Hoang Phuc\Desktop\DangBV\data.xlsx";
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
                string postTitleInput = worksheet.Cells[row, 3].Value?.ToString();
                string postContentInput = worksheet.Cells[row, 4].Value?.ToString();


                if (string.IsNullOrEmpty(emailInput) || string.IsNullOrEmpty(passwordInput) || string.IsNullOrEmpty(postTitleInput) || string.IsNullOrEmpty(postContentInput) )
                {
                    Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                    continue; // Skip to the next row
                }

                // Test registration with the data from Excel
                PostStatus(emailInput, passwordInput, postTitleInput, postContentInput);

                bool isErrorMessageDisplayed = true;
                try
                {
                    IWebElement errorMessage1 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors text-danger'] ul li"));
                    IWebElement errorMessage2 = driver.FindElement(By.CssSelector("#TenBaiViet-error"));
                    IWebElement errorMessage3 = driver.FindElement(By.CssSelector(".field-validation-error.text-danger[data-valmsg-for='NoiDung']"));
                    if (errorMessage1.Text == "Tên bài viết đã tồn tại" || errorMessage2.Text == "Bạn chưa nhập tên bài viết" || errorMessage3.Text == "Bạn chưa nhập nội dung\r\n")
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
        string filePath = @"C:\Users\Hoang Phuc\Desktop\DangBV\data.xlsx";
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                string email = worksheet.Cells[row, 1].Value.ToString();
                string password = worksheet.Cells[row, 2].Value.ToString();
                string postTitle = worksheet.Cells[row, 3].Value.ToString();
                string postContent = worksheet.Cells[row, 4].Value.ToString();
                yield return new string[] { email, password, postTitle};
            }
        }
    }
    // Update the Excel file with the result
    private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
    {

        worksheet.Cells[row, 5].Value = result;
        
    }
    private void PostStatus( string email, string password , string postTitle, string postContent)
        {
        // Your test logic using data from Excel
        driver.Navigate().GoToUrl("http://localhost:62536/nha-tuyen-dung");
        driver.FindElement(By.Id("Email")).Click();
        driver.FindElement(By.Id("Email")).SendKeys(email);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Password")).Click();
        driver.FindElement(By.Id("Password")).SendKeys(password);
        Thread.Sleep(3000);

        driver.FindElement(By.Id("Password")).SendKeys(Keys.Enter);
        Thread.Sleep(3000);

        driver.FindElement(By.CssSelector("body > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(5) > a:nth-child(1)")).Click();
        Thread.Sleep(3000);

        IWebElement post = driver.FindElement(By.XPath("//a[contains(text(),'Đăng bài viết')]"));
        post.Click();
        Thread.Sleep(3000);

        driver.FindElement(By.XPath("//textarea[@id='TenBaiViet']")).Click();
        Thread.Sleep(3000);
        driver.FindElement(By.XPath("//textarea[@id='TenBaiViet']")).SendKeys(postTitle);
        Thread.Sleep(3000);
        driver.SwitchTo().Frame(0);
        Thread.Sleep(3000);
        driver.FindElement(By.CssSelector("body p")).Click();
        {
            var element = driver.FindElement(By.CssSelector(".cke_editable"));
            js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + postContent + "'}", element);
        }
        Thread.Sleep(3000);
        driver.SwitchTo().DefaultContent();
        Thread.Sleep(3000);
        driver.FindElement(By.CssSelector("input[value='Tạo mới']")).Click();

        // You may add assertions here to verify the success of posting
    }
}
