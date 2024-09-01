using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V120.DOM;
using OpenQA.Selenium.Edge;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;

namespace LoginUser
{
    public class Tests
    {
        public IWebDriver driver = new EdgeDriver();

        [SetUp]
        public void Setup()
        {

        }
        
        [Test]
        [TestCaseSource(nameof(GetLoginCredentialsFromExcel))]
        public void TestLogin(string username, string password)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/");
            IWebElement login = driver.FindElement(By.CssSelector("#btn-login"));
            login.Click();
            Thread.Sleep(2000);

            IWebElement usernameInput = driver.FindElement(By.CssSelector("#login_email"));
            IWebElement passwordInput = driver.FindElement(By.CssSelector("#login_password"));

            usernameInput.SendKeys(username);
            passwordInput.SendKeys(password);

            Thread.Sleep(2000);

            IWebElement loginButton = driver.FindElement(By.CssSelector("#form_btn-login"));
            loginButton.Click();

            Thread.Sleep(3000);


            bool iserror = true;
            string result;
            try
            {
                IWebElement errorMessage = driver.FindElement(By.XPath("//h4[contains(text(),'Chào mừng bạn trở lại')]"));
                if (errorMessage.Text == "Chào mừng bạn trở lại")
                {
                    iserror = false;
                }
                else
                {
                    iserror = true;
                }
            }
            catch (NoSuchElementException)
            {

            }
            result = iserror? "Pass" : "Fail";
            UpdateExcelResult(username, password, result);
        }

        private static IEnumerable<string[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"C:\Users\Hoang Phuc\Desktop\LoginUser\Excel Data\data.xlsx";
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\LoginUser\Excel Data\data.xlsx";
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