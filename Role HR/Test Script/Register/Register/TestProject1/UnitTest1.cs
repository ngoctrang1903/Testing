using System;
using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;

[TestFixture]
public class DangkiTest
{
    private IWebDriver driver;
    public IDictionary<string, object> vars { get; private set; }
    private IJavaScriptExecutor js;

    [SetUp]
    public void SetUp()
    {
        driver = new ChromeDriver();
        js = (IJavaScriptExecutor)driver;
        vars = new Dictionary<string, object>();
    }

    [TearDown]
    protected void TearDown()
    {
        driver.Quit();
    }

    //[Test]
    ////public void DangkiWithExcelData()
    ////{
    ////    // Load Excel file
    ////    string excelFilePath = @"C:\Users\Hoang Phuc\Desktop\Register\data.xlsx";
    ////    FileInfo fileInfo = new FileInfo(excelFilePath);
    ////    ExcelPackage package = new ExcelPackage(fileInfo);
    ////    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

    ////    // Start from the second row (assuming the first row is headers)
    ////    int rowCount = worksheet.Dimension.Rows;
    ////    for (int row = 2; row <= rowCount; row++)
    ////    {
    ////        string companyName = worksheet.Cells[row, 1].Value.ToString();
    ////        string email = worksheet.Cells[row, 2].Value.ToString();
    ////        string phone = worksheet.Cells[row, 3].Value.ToString();
    ////        string name =worksheet.Cells[row, 4].Value.ToString();
    ////        string address = worksheet.Cells[row, 5].Value.ToString();
    ////        string password = worksheet.Cells[row, 6].Value.ToString();

    ////        // Test registration with the data from Excel
    ////        RegisterUser(companyName, email, phone, name, address, password);
    ////    }
    ////}
    [Test]
    public void DangkiWithExcelData()
    {
        // Load Excel file
        string excelFilePath = "C:\\Users\\Hoang Phuc\\Desktop\\Register\\data.xlsx";
        FileInfo fileInfo = new FileInfo(excelFilePath);

        if (!fileInfo.Exists)
        {
            Console.WriteLine("Excel file does not exist.");
            return;
        }

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                Console.WriteLine("No worksheet found in the Excel file.");
                return;
            }

            // Start from the second row (assuming the first row is headers)
            int rowCount = worksheet.Dimension.Rows;
            for (int row = 2; row <= rowCount; row++)
            {
                string companyName = worksheet.Cells[row, 1]?.Value?.ToString();
                string email = worksheet.Cells[row, 2]?.Value?.ToString();
                string phone = worksheet.Cells[row, 3]?.Value?.ToString();
                string name = worksheet.Cells[row, 4]?.Value?.ToString();
                string address = worksheet.Cells[row, 5]?.Value?.ToString();
                string password = worksheet.Cells[row, 6]?.Value?.ToString();

                if (string.IsNullOrEmpty(companyName) || string.IsNullOrEmpty(email) || string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(address) || string.IsNullOrEmpty(password))
                {
                    Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                    continue; // Skip to the next row
                }

                // Test registration with the data from Excel
                RegisterUser(companyName, email, phone, name, address, password);

                bool isErrorMessageDisplayed = true;
                //try
                //{
                //    IWebElement errorMessage1 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors error text-danger'] ul li"));
                //    IWebElement errorMessage2 = driver.FindElement(By.CssSelector("#Password-error"));
                //    if (errorMessage1.Text == "Email này đã tồn tại" || errorMessage2.Text == "Bạn chưa nhập mật khẩu")
                //    {
                //        isErrorMessageDisplayed = false;
                //    }
                //    else
                //    {
                //        isErrorMessageDisplayed = true;
                //    }
                //}
                //catch (NoSuchElementException)
                //{

                //}


                string result = isErrorMessageDisplayed ? "Pass" : "Fail";


                UpdateExcelResult(worksheet, row, result);

                package.Save();
            }

        }
        
    }


    // Assuming you have a method to check for error messages displayed during registration
    private bool CheckForErrorMessages()
    {
        bool isErrorMessageDisplayed = true;
        try
        {
            IWebElement errorMessage1 = driver.FindElement(By.CssSelector("div[class='validation-summary-errors error text-danger'] ul li"));
            IWebElement errorMessage2 = driver.FindElement(By.CssSelector("#Password-error"));
            if (errorMessage1.Text == "Email này đã tồn tại" || errorMessage2.Text == "Bạn chưa nhập mật khẩu")
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

        return false; // Placeholder
    }

    // Update the Excel file with the result
    private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
    {

        worksheet.Cells[row, 7].Value = result; 
    }


    private void RegisterUser(string companyName, string email, string phone, string name, string address, string password)
    {
        driver.Navigate().GoToUrl("http://localhost:62536/nha-tuyen-dung/Login/Register");
        driver.FindElement(By.Id("Company_Name")).Click();
        driver.FindElement(By.Id("Company_Name")).SendKeys(companyName);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Email")).Click();
        driver.FindElement(By.Id("Email")).SendKeys(email);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Phone")).Click();
        driver.FindElement(By.Id("Phone")).SendKeys(phone);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Name")).Click();
        driver.FindElement(By.Id("Name")).SendKeys(name);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Address")).Click();
        driver.FindElement(By.Id("Address")).SendKeys(address);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Password")).Click();
        driver.FindElement(By.Id("Password")).SendKeys(password);
        Thread.Sleep(3000);
        driver.FindElement(By.Id("Pasword_Confirm")).Click();
        driver.FindElement(By.Id("Pasword_Confirm")).SendKeys(password);
        Thread.Sleep(3000);

        // Wait until the register button becomes clickable
        IWebElement registerBtn = driver.FindElement(By.CssSelector("button[type='submit']"));

        registerBtn.Click();

        Thread.Sleep(5000);
    }

}
