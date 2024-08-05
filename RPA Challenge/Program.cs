using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data;
using System.Data.OleDb;
using System.Runtime.Versioning;

namespace RPA_Challenge
{
    class Program
    {
        [SupportedOSPlatform("windows")]
        public static DataTable GetDataTableFromExcel(string path)
        {
            DataTable? result = new(); 

            try
            {
                using OleDbConnection connection = new(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;HDR=YES'");
                connection.Open();

                using OleDbCommand cmd = new("SELECT * FROM [Sheet1$]", connection);
                using OleDbDataAdapter adapter = new(cmd);
                adapter.Fill(result);
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred while reading Excel file: " + e.Message);
            }

            return result;
        }

        static void Main(string[] args)
        {
            string path = @"C:\challenge.xlsx";
            Program p = new();

            if (!OperatingSystem.IsWindows())
            {
                Console.WriteLine("This application is only supported on Windows.");
                return;
            }

            DataTable extractedDT = GetDataTableFromExcel(path);

            DataTable filteredDT = extractedDT.AsEnumerable()
               .Where(row => row.ItemArray.Any(field => field != DBNull.Value && !string.IsNullOrWhiteSpace(field?.ToString())))
               .CopyToDataTable();

            Console.WriteLine($"Number of rows after filtering: {filteredDT.Rows.Count}");


            if (filteredDT.Rows.Count == 0)
            {
                Console.WriteLine("No data found in Excel file.");
                return;
            }

            using ChromeDriver driver = new(Environment.CurrentDirectory);
            try
            {
                driver.Navigate().GoToUrl("http://www.rpachallenge.com/");
                driver.Manage().Window.Maximize();

                driver.FindElement(By.XPath("//a[text()='Input Forms']")).Click();
                driver.FindElement(By.XPath("//button[text()='Start']")).Click();

                foreach (DataRow row in filteredDT.Rows)
                {

                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelFirstName']")).SendKeys(row[0].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelLastName']")).SendKeys(row[1].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelCompanyName']")).SendKeys(row[2].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelRole']")).SendKeys(row[3].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelAddress']")).SendKeys(row[4].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelEmail']")).SendKeys(row[5].ToString());
                    driver.FindElement(By.XPath("//input[@ng-reflect-name='labelPhone']")).SendKeys(row[6].ToString());

                    driver.FindElement(By.XPath("//input[@value='Submit']")).Click();
                }

                Console.WriteLine("RPA Challenge completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the RPA challenge: " + ex.Message);
            }
            finally
            {
                driver.Quit();
            }
        }
    }
}
