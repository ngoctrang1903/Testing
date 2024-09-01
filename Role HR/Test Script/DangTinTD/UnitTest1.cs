using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading;
using System.Net.Mail;

namespace DangTinTD
{
    public class Tests
    {
        public IWebDriver driver;

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
        public void EditWithExcelData()
        {
            string filePath = @"C:\Users\Hoang Phuc\Desktop\DangTinTD\ExcelData\data.xlsx";
            FileInfo file = new FileInfo(filePath);
            if(!file.Exists )
            {
                Console.WriteLine("Excel file does not exist.");
                return;
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                {
                    Console.WriteLine("No worksheet found in the Excel file.");
                    return;
                }
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string user = worksheet.Cells[row, 1].Value.ToString();
                    string pass = worksheet.Cells[row, 2].Value.ToString();
                    string jobname = worksheet.Cells[row, 3].Value.ToString();
                    string jobrank = worksheet.Cells[row, 4].Value.ToString();
                    string jobtype = worksheet.Cells[row, 5].Value.ToString();
                    string salary = worksheet.Cells[row, 6].Value.ToString();
                    string major = worksheet.Cells[row, 7].Value.ToString();
                    string jobaddress = worksheet.Cells[row, 8].Value.ToString();
                    string sex = worksheet.Cells[row, 9].Value.ToString();
                    string homeaddress = worksheet.Cells[row, 10].Value.ToString();
                    string description = worksheet.Cells[row, 11].Value.ToString();
                    string require = worksheet.Cells[row, 12].Value.ToString();
                    string skill = worksheet.Cells[row, 13].Value.ToString();
                    string quyenloinhanduoc = worksheet.Cells[row, 14].Value.ToString();
                    string num = worksheet.Cells[row, 15].Value.ToString();
                    string time = worksheet.Cells[row, 16].Value.ToString();

                    if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(jobname) || string.IsNullOrEmpty(jobrank) || string.IsNullOrEmpty(jobtype) || string.IsNullOrEmpty(salary) || string.IsNullOrEmpty(major) || string.IsNullOrEmpty(jobaddress) || string.IsNullOrEmpty(sex) || string.IsNullOrEmpty(homeaddress) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(require) || string.IsNullOrEmpty(skill) || string.IsNullOrEmpty(quyenloinhanduoc) || string.IsNullOrEmpty(num) || string.IsNullOrEmpty(time))
                    {
                        Console.WriteLine($"One or more fields are empty in row {row}. Skipping registration for this row.");
                        continue;
                    }
                    PostTTD(user, pass, jobname, jobrank, jobtype, salary, major, jobaddress, sex, homeaddress, description, require, skill, quyenloinhanduoc, num, time);
                    
                    bool isErrorMessageDisplayed = false;

                    try
                    {
                        IWebElement ketqua = driver.FindElement(By.XPath("//a[contains(text(),'Sửa')]"));
                        ketqua.Click();

                        isErrorMessageDisplayed = true;
                        if (ketqua == null)
                        {
                            isErrorMessageDisplayed = true;
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
            string filePath = @"C:\Users\Hoang Phuc\Desktop\DangTinTD\ExcelData\data.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string username = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string tencv = worksheet.Cells[row, 3].Value.ToString();
                    string baccv = worksheet.Cells[row, 4].Value.ToString();
                    string loaicv = worksheet.Cells[row, 5].Value.ToString();
                    string mucluong = worksheet.Cells[row, 6].Value.ToString();
                    string chuyennganh = worksheet.Cells[row, 7].Value.ToString();
                    string diachi = worksheet.Cells[row, 8].Value.ToString();
                    string gioitinh = worksheet.Cells[row, 9].Value.ToString();
                    string address = worksheet.Cells[row, 10].Value.ToString();
                    string mota = worksheet.Cells[row, 11].Value.ToString();
                    string yeucau = worksheet.Cells[row, 12].Value.ToString();
                    string kynang = worksheet.Cells[row, 13].Value.ToString();
                    string quyenloi = worksheet.Cells[row, 14].Value.ToString();
                    string soluong = worksheet.Cells[row, 15].Value.ToString();
                    string thoigian = worksheet.Cells[row, 16].Value.ToString();

                    yield return new string[] { username, password, tencv, baccv, loaicv, mucluong, chuyennganh, diachi, gioitinh, address, mota, yeucau, kynang, quyenloi, soluong, thoigian };
                }
            }
        }
        private void UpdateExcelResult(ExcelWorksheet worksheet, int row, string result)
        {
            worksheet.Cells[row, 17].Value = result;

        }
        public void PostTTD(string username, string password, string tencv, string baccv, string loaicv, string mucluong, string chuyennganh, string diachi, string gioitinh, string address, string mota, string yeucau, string kynang, string quyenloi, string soluong, string thoigian)
        {
            driver.Navigate().GoToUrl("http://localhost:62536/nha-tuyen-dung");

            IWebElement usernameInput = driver.FindElement(By.CssSelector("#Email"));
            usernameInput.SendKeys(username);
            IWebElement passwordInput = driver.FindElement(By.CssSelector("#Password"));
            passwordInput.SendKeys(password);
            Thread.Sleep(5000);
            driver.FindElement(By.CssSelector("button[type='submit']")).Click();

            driver.FindElement(By.CssSelector("body > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(3) > a:nth-child(1) > span:nth-child(2)")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/TinTuyenDung/Create']")).Click();
            Thread.Sleep(2000);

            IWebElement TenCV = driver.FindElement(By.CssSelector("#TenCongViec"));
            TenCV.SendKeys(tencv);
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn cấp bậc --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaCapBac"));
                dropdown.FindElement(By.XPath("//option[. = '" + baccv + "']")).Click();
            }
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn hình thức làm việc --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaLoaiCV"));
                dropdown.FindElement(By.XPath("//option[. = '" + loaicv + "']")).Click();
            }
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn mức lương --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaMucLuong"));
                dropdown.FindElement(By.XPath("//option[. = '" + mucluong + "']")).Click();
            }
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn chuyên ngành --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaCN"));
                dropdown.FindElement(By.XPath("//option[. = '" + chuyennganh + "']")).Click();
            }
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn địa chỉ --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("MaDiaChi"));
                dropdown.FindElement(By.XPath("//option[. = '" + diachi + "']")).Click();
            }
            Thread.Sleep(2000);
            //driver.FindElement(By.CssSelector("button[title='-- Chọn giới tính yêu cầu --']")).Click();
            {
                var dropdown = driver.FindElement(By.Id("GioiTinhYC"));
                dropdown.FindElement(By.XPath("//option[. = '" + gioitinh + "']")).Click();
            }
            Thread.Sleep(2000);
            IWebElement diachilamviec = driver.FindElement(By.CssSelector("#DiaChiLamViec"));
            diachilamviec.SendKeys(address);
            Thread.Sleep(2000);

            //Mo ta cong viec
            driver.SwitchTo().Frame(0);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + mota + "'}", element);
            }
            Thread.Sleep(2000);
            driver.SwitchTo().DefaultContent();
            //Yeu cau ung vien
            driver.SwitchTo().Frame(1);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + yeucau + "'}", element);
            }
            Thread.Sleep(2000);
            driver.SwitchTo().DefaultContent();
            //Ky nang lien quan
            driver.SwitchTo().Frame(2);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + kynang + "'}", element);
            }
            Thread.Sleep(2000);
            driver.SwitchTo().DefaultContent();
            //Quyen loi
            driver.SwitchTo().Frame(3);
            driver.FindElement(By.CssSelector("html")).Click();
            {
                var element = driver.FindElement(By.CssSelector(".cke_editable"));
                js.ExecuteScript("if(arguments[0].contentEditable === 'true') {arguments[0].innerText = '" + quyenloi + "'}", element);
            }
            Thread.Sleep(2000);
            driver.SwitchTo().DefaultContent();

            IWebElement SoluongTuyen = driver.FindElement(By.CssSelector("#SoLuong"));
            SoluongTuyen.SendKeys(soluong);
            Thread.Sleep(1000);
            IWebElement HanNop = driver.FindElement(By.XPath("//input[@id='HanNop']"));
            HanNop.SendKeys(thoigian);
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector(".mt-3 > .col-md-6")).Click();
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
        }
        
    }
}