using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;
using System.Reflection;
using log4net;
using log4net.Config;

class Program
{
    // Initialize log4net logger
    private static readonly ILog log = LogManager.GetLogger(typeof(Program));

    static void Main(string[] args)
    {
        // Configure log4net
        ConfigureLogging();

        log.Info("Starting scraper...");

        var options = new ChromeOptions();
        options.AddArguments("--headless=new", "--disable-gpu", "--window-size=1920,1080", "--blink-settings=imagesEnabled=false");

        using var driver = new ChromeDriver(options);
        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        try
        {
            driver.Navigate().GoToUrl("https://conversion.straive.com/eConnect2/Main.aspx");
            log.Info("Navigated to login page.");

            driver.FindElement(By.Id("txtUser")).SendKeys("vietnam");
            driver.FindElement(By.Id("txtPass")).SendKeys("1234567");
            driver.FindElement(By.Id("imgOK")).Click();
            log.Info("Login submitted.");

            wait.Until(d => d.Url.Contains("Main.aspx"));
            wait.Until(d => d.FindElement(By.Id("lblCountJobs")));

            int totalJobs = GetTotalJobs(driver.PageSource);
            int totalPages = CalculateTotalPages(totalJobs, 100);
            log.Info($"Total jobs: {totalJobs}, Total pages: {totalPages}");

            List<JobRecord> allRows = new List<JobRecord>();
            string outputFileName = $"Scraper_{DateTime.Now:yyyyMMdd_HHmmss}.html";

            using (StreamWriter sw = new StreamWriter(outputFileName, true))
            {
                sw.WriteLine("<html><body>");

                for (int i = 1; i <= totalPages; i++)
                {
                    log.Info($"Scraping page {i}...");

                    // Wait for table to be present and visible
                    var table = wait.Until(d => {
                        var element = d.FindElement(By.XPath("//table[@id='GridData_New']"));
                        return element.Displayed ? element : null;
                    });

                    if (table == null)
                    {
                        log.Warn($"Table not found on page {i}. Retrying...");
                        driver.Navigate().Refresh();
                        Thread.Sleep(3000);
                        table = wait.Until(d => d.FindElement(By.XPath("//table[@id='GridData_New']")));
                    }

                    string currentTableHtml = table.GetAttribute("outerHTML");

                    // Verify that the table contains data
                    var rows = table.FindElements(By.TagName("tr")).ToList();
                    int rowCount = rows.Count - 1; // Subtract header row

                    if (rowCount <= 0)
                    {
                        log.Warn($"No data rows found on page {i}");
                        continue;
                    }

                    // Extract the data to verify it's correct
                    List<JobRecord> pageRecords = ExtractTableData(table);
                    log.Info($"Extracted {pageRecords.Count} records from page {i}");

                    // Write the table HTML to the file with a page identifier
                    sw.WriteLine($"<!-- Page {i} data -->");
                    sw.WriteLine(currentTableHtml);
                    sw.Flush(); // Force write to disk

                    allRows.AddRange(pageRecords);
                    log.Info($"Successfully saved table from page {i} with {pageRecords.Count} records");

                    // Move to next page if not on the last page
                    if (i < totalPages)
                    {
                        string nextPage = (i + 1).ToString();

                        try
                        {
                            var nextPageLink = driver.FindElement(By.XPath($"//a[contains(@href, '__doPostBack') and text()='{nextPage}']"));

                            // Scroll to make sure the element is in view
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", nextPageLink);
                            Thread.Sleep(500);

                            nextPageLink.Click();

                            // Wait for page to load by checking for a changing element
                            wait.Until(d => {
                                try
                                {
                                    string currentPageIndicator = d.FindElement(By.CssSelector(".dgPager span")).Text;
                                    return currentPageIndicator == nextPage;
                                }
                                catch
                                {
                                    return false;
                                }
                            });

                            Thread.Sleep(2000); // Additional wait to ensure complete loading

                            log.Info($"Moved to page {nextPage}.");
                        }
                        catch (Exception ex)
                        {
                            log.Error($"Error navigating to page {nextPage}", ex);
                            break;
                        }
                    }
                }

                sw.WriteLine("</body></html>");
            }

            log.Info($"Scraping completed. Total records: {allRows.Count}");
            log.Info($"Output saved to {outputFileName}");

            // Additionally, you could save the structured data to CSV or another format
            SaveToCsv(allRows, $"JobRecords_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
        }
        catch (Exception ex)
        {
            log.Error("Critical error in scraping process", ex);
        }

        driver.Quit();
        log.Info("Browser closed.");
    }

    static void ConfigureLogging()
    {
        // Create a configuration for log4net
        var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly()!);

        // Create and configure appenders programmatically
        var fileAppender = new log4net.Appender.RollingFileAppender
        {
            Name = "RollingFileAppender",
            File = "logs/scraper.log",
            AppendToFile = true,
            RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Date,
            DatePattern = "yyyy-MM-dd",
            LockingModel = new log4net.Appender.FileAppender.MinimalLock(),
            Layout = new log4net.Layout.PatternLayout("%date [%thread] %-5level %logger - %message%newline")
        };
        fileAppender.ActivateOptions();

        var consoleAppender = new log4net.Appender.ConsoleAppender
        {
            Name = "ConsoleAppender",
            Layout = new log4net.Layout.PatternLayout("%date [%thread] %-5level - %message%newline")
        };
        consoleAppender.ActivateOptions();

        // Configure the root logger
        var hierarchy = (log4net.Repository.Hierarchy.Hierarchy)logRepository;
        hierarchy.Root.AddAppender(fileAppender);
        hierarchy.Root.AddAppender(consoleAppender);
        hierarchy.Root.Level = log4net.Core.Level.Info;
        hierarchy.Configured = true;
    }

    static List<JobRecord> ExtractTableData(IWebElement table)
    {
        List<JobRecord> records = new List<JobRecord>();
        var rows = table.FindElements(By.TagName("tr")).ToList();

        for (int i = 1; i < rows.Count; i++) // Skip header row
        {
            try
            {
                var cells = rows[i].FindElements(By.TagName("td"));
                if (cells.Count < 15) continue;

                var job = new JobRecord
                {
                    JobID = cells[0].Text.Trim(),
                    ChildID = cells[1].Text.Trim(),
                    JobType = cells[2].Text.Trim(),
                    WorkCode = cells[3].Text.Trim(),
                    Characters = cells[4].Text.Trim(),
                    FilesOrPages = cells[5].Text.Trim(),
                    ClientDueDate = cells[6].Text.Trim(),
                    AssignDate = cells[7].Text.Trim(),
                    DateAccepted = cells[8].Text.Trim(),
                    ReturnDate = cells[9].Text.Trim(),
                    ActualReturnDate = cells[10].Text.Trim(),
                    Attachment = cells[11].Text.Trim(),
                    Remarks = cells[12].Text.Trim(),
                    Shipment = cells[13].Text.Trim(),
                    QueryStatus = cells[14].Text.Trim()
                };

                records.Add(job);
            }
            catch (Exception ex)
            {
                log.Error($"Error extracting row {i}", ex);
            }
        }

        return records;
    }

    static void SaveToCsv(List<JobRecord> records, string fileName)
    {
        try
        {
            using (var writer = new StreamWriter(fileName))
            {
                // Write header
                writer.WriteLine("JobID,ChildID,JobType,WorkCode,Characters,FilesOrPages,ClientDueDate,AssignDate,DateAccepted,ReturnDate,ActualReturnDate,Attachment,Remarks,Shipment,QueryStatus");

                // Write data rows
                foreach (var record in records)
                {
                    writer.WriteLine($"{CsvEscape(record.JobID)}," +
                                    $"{CsvEscape(record.ChildID)}," +
                                    $"{CsvEscape(record.JobType)}," +
                                    $"{CsvEscape(record.WorkCode)}," +
                                    $"{CsvEscape(record.Characters)}," +
                                    $"{CsvEscape(record.FilesOrPages)}," +
                                    $"{CsvEscape(record.ClientDueDate)}," +
                                    $"{CsvEscape(record.AssignDate)}," +
                                    $"{CsvEscape(record.DateAccepted)}," +
                                    $"{CsvEscape(record.ReturnDate)}," +
                                    $"{CsvEscape(record.ActualReturnDate)}," +
                                    $"{CsvEscape(record.Attachment)}," +
                                    $"{CsvEscape(record.Remarks)}," +
                                    $"{CsvEscape(record.Shipment)}," +
                                    $"{CsvEscape(record.QueryStatus)}");
                }
            }
            log.Info($"CSV data saved to {fileName}");
        }
        catch (Exception ex)
        {
            log.Error($"Error saving CSV", ex);
        }
    }

    static string CsvEscape(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }
        return value;
    }

    static int GetTotalJobs(string html)
    {
        var match = Regex.Match(html, @"No\. of jobs: <b>(\d+)</b>");
        return match.Success ? int.Parse(match.Groups[1].Value) : 0;
    }

    static int CalculateTotalPages(int totalRecords, int perPage)
    {
        return (int)Math.Ceiling((double)totalRecords / perPage);
    }
}

public class JobRecord
{
    public string? JobID { get; set; }
    public string? ChildID { get; set; }
    public string? JobType { get; set; }
    public string? WorkCode { get; set; }
    public string? Characters { get; set; }
    public string? FilesOrPages { get; set; }
    public string? ClientDueDate { get; set; }
    public string? AssignDate { get; set; }
    public string? DateAccepted { get; set; }
    public string? ReturnDate { get; set; }
    public string? ActualReturnDate { get; set; }
    public string? Attachment { get; set; }
    public string? Remarks { get; set; }
    public string? Shipment { get; set; }
    public string? QueryStatus { get; set; }
}