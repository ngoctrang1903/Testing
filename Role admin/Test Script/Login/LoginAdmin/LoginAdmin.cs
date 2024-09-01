using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace TestProject1
{
    public class Tests
    {
        public IWebDriver driver = new EdgeDriver();

        [SetUp]
        public void Setup()
        {
        }

        [Test]
        [TestCase("http://localhost:62536/Admin/Login")]
        public void Test1OpenPortal(string url)
        {
            driver.Navigate().GoToUrl(url);

            Assert.Pass();
        }

        [Test]
        [TestCase("UserName", "admin@gmail.com")]
        public void Test2InsertName(string Name, string content)
        {
            IWebElement NameInput = driver.FindElement(By.Name(Name));

            if (NameInput != null)
            {
                NameInput.SendKeys(content);
            }
            Thread.Sleep(2000);

            Assert.Pass();
        }

        [Test]
        [TestCase("PassWord", "12345678a")]
        public void Test3InsertPass(string PassName, string content)
        {
            IWebElement PassInput = driver.FindElement(By.Name(PassName));

            if (PassInput != null)
            {
                PassInput.SendKeys(content);
            }
            Thread.Sleep(2000);

            Assert.Pass();
        }

        [Test]
        [TestCaseSource(nameof(GetLoginCredentialsFromExcel))]
        public void TestLogin(string username, string password)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/Admin/Login");

            IWebElement usernameInput = driver.FindElement(By.Name("UserName"));
            IWebElement passwordInput = driver.FindElement(By.Name("PassWord"));

            usernameInput.SendKeys(username);
            passwordInput.SendKeys(password);

            Thread.Sleep(2000);

            IWebElement loginButton = driver.FindElement(By.XPath("//button[contains(text(),'Đăng nhập')]"));
            loginButton.Click();

            Thread.Sleep(3000);


            bool isErrorMessageDisplayed = true;
            try
            {
                IWebElement errorMessage = driver.FindElement(By.CssSelector("div[class='validation-summary-errors error text-danger'] ul li"));
                if (errorMessage.Text == "Sai tài khoản hoặc mật khẩu")
                {
                    isErrorMessageDisplayed = false;
                }
                else
                {
                    isErrorMessageDisplayed = true;
                }
            }
            catch (NoSuchElementException)
            {

            }


            string result = isErrorMessageDisplayed ? "Pass" : "Fail";


            UpdateExcelResult(username, password, result);
        }

        private static IEnumerable<string[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"C:\Users\Admin\Desktop\Login\LoginAdmin.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string username = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    yield return new string[] { username, password };
                }
            }
        }

        private void UpdateExcelResult(string username, string password, string result)
        {
            string filePath = @"C:\Users\Admin\Desktop\Login\LoginAdmin.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string user = worksheet.Cells[row, 1].Value.ToString();
                    string pass = worksheet.Cells[row, 2].Value.ToString();

                    if (user == username && pass == password)
                    {
                        worksheet.Cells[row, 3].Value = result;
                        break;
                    }
                }

                package.Save();
            }
        }
    }
}
