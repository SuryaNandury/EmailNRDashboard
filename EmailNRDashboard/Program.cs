using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Configuration;
using System.Globalization;


namespace EmailNRDashboard
{
    internal class Program
    {
        private static string _processLog;
        private static IWebDriver _driver;
        private static WebDriverWait _wait;
        private static string _currentDirectory;
        private static string _rptTitle;

        private static void Main()
        {
            _rptTitle = "GPI Website Metrics from " + DateTime.Now.AddDays(-30).ToString("MMM dd, yyyy") + " to " + DateTime.Now.ToString("MMM dd, yyyy");
            _currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            _processLog = _rptTitle + "<br/><br/>\n\n";
            _processLog += _currentDirectory+ "<br/><br/>\n\nDeleting old pdf files (if any)<br/><br/>\n\n";
            var oldfiles = Directory.GetFiles(_currentDirectory, "*.pdf");
            foreach (var oldfile in oldfiles)
            {
                File.Delete(oldfile);
            }
            _processLog = _rptTitle;

            _processLog += "Process started at " + DateTime.UtcNow.ToString(CultureInfo.InvariantCulture) + "<br/><br/>\n\n";
            ChromeOptions chromeOptions = new ChromeOptions();
            try
            {
                chromeOptions.AddUserProfilePreference("download.default_directory", _currentDirectory);
                chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                chromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
                chromeOptions.AddUserProfilePreference("plugins.plugins_disabled", "Chrome PDF Viewer");
                chromeOptions.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
                _driver = new ChromeDriver(chromeOptions);
                _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
            }
            catch (Exception ex)
            {
                _processLog += "Driver declaration failed. " + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                _driver.Quit();
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
            _driver.Navigate().GoToUrl("https://one.newrelic.com/");
            _processLog += "Opened https://one.newrelic.com/ <br/><br/>\n\n";
            if (_driver.Url.Contains("https://adfs.graphicpkg.com/"))
            {
                HandleAdFs(_driver);
            }
            IWebElement loginEmail = _driver.FindElement(By.Id("login_email"));
            loginEmail.SendKeys("Surya.Nandury@GraphicPkg.com");
            IWebElement loginPassword = _driver.FindElement(By.Id("login_password"));
            loginPassword.SendKeys("Random@0707" + Keys.Enter);
            _processLog += "Logged in successfully into https://one.newrelic.com/" + "<br/><br/>\n\n";
            if (_driver.Url.Contains("https://adfs.graphicpkg.com/"))
            {
                HandleAdFs(_driver);
            }
            IWebElement dashboardElement;
            try
            {
                dashboardElement = _wait.Until(condition: ExpectedConditions.ElementIsVisible(By.XPath("//p[@class='nr1-LauncherButton-name' and text()='Dashboards']")));
            }
            catch (Exception ex)
            {
                _processLog += "Dashboard element could not be found or isn't loaded." + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                _driver.Quit();
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
            dashboardElement.Click();
            _processLog += "Clicked on dashboard." + "<br/><br/>\n\n";
            IWebElement dashboard1;
            IWebElement dashboard2;
            try
            {
                dashboard1 = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@class='NameCell-name' and text()='Executive Dashboard - Page 1']")));
            }
            catch (Exception ex)
            {
                _processLog += "Executive Dashboard - Page 1 could not be found." + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                _driver.Quit();
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
            dashboard1.Click();
            _processLog += "Clicked on Executive Dashboard - Page 1. Waiting for download button to be visible." + "<br/><br/>\n\n";
            Thread.Sleep(5000);
            IWebElement downloadbutton1 = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div/div[3]/div[1]/div[1]/div[2]/div[2]/button")));

            downloadbutton1.Click();
            _processLog += "Clicked on export button 1 after waiting for 5 seconds." + "<br/><br/>\n\n";
            Thread.Sleep(5000);
            _processLog += "Waiting for 5 seconds before clicking on back button." + "<br/><br/>\n\n";
            _driver.Navigate().Back();
            Thread.Sleep(5000);

            try
            {
                dashboard2 = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@class='NameCell-name' and text()='Executive Dashboard - Page 2']")));
            }
            catch (Exception ex)
            {
                _processLog += "Executive Dashboard - Page 2 could not be found." + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                _driver.Quit();
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
            dashboard2.Click();
            _processLog += "Clicked on Executive Dashboard - Page 2. Waiting for download button to be visible." + "<br/><br/>\n\n";
            Thread.Sleep(5000);
            IWebElement downloadbutton2 = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div/div[3]/div[1]/div[1]/div[2]/div[2]/button")));

            downloadbutton2.Click();
            _processLog += "Clicked on export button 2 after waiting for 5 seconds." + "<br/><br/>\n\n";
            Thread.Sleep(5000);

            string[] downloadedFiles = { "Executive Dashboard - Page 1.pdf.pdf", "Executive Dashboard - Page 2.pdf.pdf" };

            MergePdFs(_currentDirectory, downloadedFiles);

            _driver.Quit();
            WriteLog(_processLog);
            SendEmail("Report");


        }

        private static void HandleAdFs(IWebDriver driver)
        {
            if (!driver.Url.Contains("https://adfs.graphicpkg.com/")) return;
            _processLog += "Opened " + driver.Url + "<br/><br/>\n\n";
            IWebElement usernameElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("userNameInput")));
            usernameElement.Clear();
            usernameElement.SendKeys("Surya.Nandury@GraphicPkg.com");
            driver.FindElement(By.Id("passwordInput")).SendKeys("Divya@0790" + Keys.Enter);
            if (driver.FindElements(By.Id("error")).Count >= 1)
            {
                if (driver.FindElement(By.Id("errorText")).Text.ToLower().Contains("incorrect user id or password. type the correct user id and password, and try again."))
                {
                    _processLog += "Login failed " + "<br/><br/>\n\n";
                    driver.Quit();
                    SendEmail("Error");
                    WriteLog(_processLog);
                }
            }

            if (!driver.Url.Contains("https://login.zscalertwo.net/")) return;
            try
            {
                IWebElement acceptButton = driver.FindElement(By.XPath("/html/body/div/div/form/table[2]/tbody/tr/td/table[1]/tbody/tr[4]/td/table/tbody/tr/td[1]/input[2]"));
                acceptButton.Click();
                _processLog += "Zscaler accept page. Accept button clicked " + "<br/><br/>\n\n";
            }
            catch (Exception ex)
            {
                _processLog += "Login failed on Zscaler page. Accept button could not be found probably?" + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                driver.Quit();
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
        }

        private static void MergePdFs(string targetPath, params string[] pdfs)
        {
            try
            {
                using (PdfDocument targetDoc = new PdfDocument())
                {
                    foreach (string pdf in pdfs)
                    {
                        using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < pdfDoc.PageCount; i++)
                            {
                                targetDoc.AddPage(pdfDoc.Pages[i]);
                                File.Delete(targetPath + pdf);
                            }
                        }
                    }
                    targetDoc.Save(targetPath + _rptTitle + ".pdf");
                }
            }
            catch (Exception ex)
            {
                _processLog += "Failed either while combining or downloading the pdf's" + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                SendEmail("Error");
                WriteLog(_processLog);
                throw;
            }
        }

        public static void SendEmail(string type)
        {
            MailMessage mail = new MailMessage();
            SmtpClient client = new SmtpClient("gpimail.na.graphicpkg.pri");
            mail.From = new MailAddress("DoNotReply@GraphicPkg.com");
            mail.IsBodyHtml = true;
            if (!ConfigurationManager.AppSettings.AllKeys.Contains("ToAddresses"))
            {
                _processLog += "To address is not defined in config file.<br/><br/>\n\n";
                WriteLog(_processLog);
                throw new Exception("No To Address");
            }
            if (!ConfigurationManager.AppSettings.AllKeys.Contains("CcAddresses"))
            {
                _processLog += "Cc address is not defined in config file.<br/><br/>\n\n";
                WriteLog(_processLog);
                throw new Exception("No Cc Address");
            }
            string[] toaddr = ConfigurationManager.AppSettings["ToAddresses"].Replace(" ", string.Empty).Split(',');
            string[] ccaddr = ConfigurationManager.AppSettings["CcAddresses"].Replace(" ", string.Empty).Split(',');
            string[] bccaddr = { "Surya.Nandury@GraphicPkg.com", "Vara.Karyampudi@GraphicPkg.com" };

            switch (type)
            {
                case "Error":
                {
                    foreach (string item in bccaddr)
                    {
                        mail.To.Add(item);
                    }
                    mail.Subject = "Error occurred while generating " + _rptTitle;
                    mail.Body = _processLog;
                    break;
                }
                case "Report":
                {
                    foreach (string item in toaddr)
                    {
                        mail.To.Add(item);
                    }

                    if (ccaddr.Any())
                    {
                        foreach (string item in ccaddr)
                        {
                            mail.CC.Add(item);
                        }
                    }
                    foreach (string item in bccaddr)
                    {
                        mail.Bcc.Add(item);
                    }
                    mail.Subject = _rptTitle;
                    const string body = "<div style=\"font-family:Calibri; font-size:14px;\"> <p> Greetings!! </p> <p> Attached, is a report on {0}.<br /> For any questions regarding the attached report, please write to <a href=\"mailto:DLIST_Public_Web@graphicpkg.com\">DLIST_Public_Web</a> </p> <p> Regards<br /> GPI Website Support Team </p> <p style=\"color:#686868;\"> ****This is an automated email notification. Do not respond to this message.****</p> </div>";
                    mail.Body = string.Format(body, _rptTitle);
                    mail.Attachments.Add(new Attachment(_currentDirectory + _rptTitle + ".pdf"));
                    break;
                }
            }
            try
            {
                client.Send(mail);
            }
            catch (Exception ex )
            {
                _processLog += "Failed to send message or delete file." + "\nMessage\n" + ex.Message + "\nStack Trace\n" + ex.StackTrace + "<br/><br/>\n\n";
                WriteLog(_processLog);
                throw;
            }
        }

        public static void WriteLog (string data)
        {
            using (StreamWriter sw = File.CreateText(_currentDirectory +  "Log.txt"))
            {
                sw.WriteLine(data);
            }
        }
    }
}
