using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager;

// Resources:
// https://www.nuget.org/packages/Selenium.WebDriver
// https://www.selenium.dev/documentation/webdriver/getting_started/install_drivers/
// https://github.com/rosolko/WebDriverManager.Net
// https://www.newtonsoft.com/json
// https://github.com/JamesNK/Newtonsoft.Json/releases



namespace WebDriverDownloader
{

    class Program
    {

        public static void Main(string[] args)
        {

            string url = @"https://support.hp.com/us-en/help/hp-support-assistant";
            DownloaderFromHtml download = new DownloaderFromHtml();
            //download.Sel(url, "hpsabtn1");


            url = "https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html";
            //url = "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html";
            List<string> downloadLink = download.FindFile(url, "href");
            if (downloadLink.Count() > 0)
            {
                // download and install using downloadLink
                // ......
                foreach(var item in downloadLink)
                    Console.WriteLine(item);

            }
            else
                Console.WriteLine("File not found.");

        }
    }


    public class DownloaderFromHtml
    {

        public List<string> FindFile(string destUrl, string tagName)
        {
            List<string> list = new List<string>();
            var web = new HtmlWeb(); // Create new HtmlWeb instance
            try
            {
                var doc = web.Load(destUrl); // load html documnt from target url
                foreach (var link in doc.DocumentNode.Descendants("a")) // Fetch all a elements from html doc
                {
                    var href = link.GetAttributeValue(tagName, string.Empty); // Get the tag (by tagName) values
                    if (!string.IsNullOrEmpty(href) && (href.EndsWith(".exe") || href.EndsWith(".msi"))) // get valus for msi or exe files
                    {
                        list.Add(href); // Add download link to list
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return list; 
        }

        public bool Sel(string targetUrl, string elemId)
        {

            // Format time and date to be used as folder name:
            string nowTime = DateTime.Now.ToString().Replace(":", "-").Replace("/", "-");
            string sourceFolder = System.AppDomain.CurrentDomain.BaseDirectory + "source\\tempfiles\\";

            // Create oprions for web driver:
            EdgeOptions optionsEdge = new EdgeOptions();
            optionsEdge.AddArguments("--disable-notifications"); // disable notifications
            optionsEdge.AddArguments("--disable-extensions"); // disable browser extensions
            optionsEdge.AddArguments("--disable-infobars"); // disable the "chrome is being controlled by automated test software" infobar
            optionsEdge.AddArguments("--headless=new"); // Run dirver headless (Browser not visiable)
            //optionsEdge.AddArguments("--disable-gpu"); // this option will hide the download bar
            optionsEdge.AddArguments("--log-level=3"); // Lower logs in console
            optionsEdge.AddArguments("--inprivate"); // run driver in private mode
            optionsEdge.AddUserProfilePreference("download.default_directory", (sourceFolder + nowTime)); // default directory for downloading files
            optionsEdge.AddUserProfilePreference("download.prompt_for_download", false); // don't prompt download location
            optionsEdge.AddUserProfilePreference("safebrowsing.enabled", "true"); // allow files to download without "harmful file message)


            // install the supported edge driver (driver that matches the current installed Edge version):
            new DriverManager().SetUpDriver(new EdgeConfig());

            // Create EdgeDriver instance
            IWebDriver driver = new EdgeDriver(optionsEdge);
            driver.Navigate().GoToUrl(targetUrl); // go to target download site (targetUrl)

            // Create folder for file download location:
            Directory.CreateDirectory(sourceFolder + nowTime);

            // Make sure folder is empty:
            foreach (var item in Directory.GetFiles((sourceFolder + nowTime + "\\"), " *.exe").ToArray())
                File.Delete(item);
            try
            {
                // Web driver - download file:
                IWebElement link = driver.FindElement(By.Id(elemId)); // find link to File by element ID (elemId)
                String javascript = "arguments[0].click()"; // click on the first argument of "link". Note: the herf value is "javascript:void(0)"
                IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
                executor.ExecuteScript(javascript, link); // Execute the javascript code to download the file
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }


            // Wait for the file to finish downloading:
            Console.WriteLine("Downloading file, Please wait...");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            wait.Until(wd =>
            {
                string[] files = Directory.GetFiles((sourceFolder + nowTime), "*.exe");
                if (files.Length == 0)
                    return false; // The file hasn't downloaded yet
                else
                    return true; // The file has downloaded
            });


            // Install exe file:
            // ...
            // ...


            driver.Quit();

            // Delete files after installation complete:
            if (Directory.Exists(sourceFolder));
                Directory.Delete(sourceFolder, true);
            if (Directory.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "Edge")) 
                Directory.Delete((System.AppDomain.CurrentDomain.BaseDirectory + "Edge"), true); // Delete Edge webdriver files
            

            return true;
        }
    }
}