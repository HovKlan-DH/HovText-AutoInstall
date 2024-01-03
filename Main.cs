using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Management;
using System.Linq;


namespace HovText_Update
{
    public partial class Main : Form
    {
        string hovtextPage = "https://hovtext.com";
        string[] startupArgs;
        int secs = 5;
        string appVer = "";

        public Main(string[] args)
        {
            InitializeComponent();
            label3.Text = "";
            startupArgs = args;

            // Subscribe to the Load event
            Load += new EventHandler(Main_Load);

            // Get application file version from assembly
            Assembly assemblyInfo = Assembly.GetExecutingAssembly();
            string assemblyVersion = FileVersionInfo.GetVersionInfo(assemblyInfo.Location).FileVersion;
            string year = assemblyVersion.Substring(0, 4);
            string month = assemblyVersion.Substring(5, 2);
            string day = assemblyVersion.Substring(8, 2);
            switch (month)
            {
                case "01": month = "January"; break;
                case "02": month = "February"; break;
                case "03": month = "March"; break;
                case "04": month = "April"; break;
                case "05": month = "May"; break;
                case "06": month = "June"; break;
                case "07": month = "July"; break;
                case "08": month = "August"; break;
                case "09": month = "September"; break;
                case "10": month = "October"; break;
                case "11": month = "November"; break;
                case "12": month = "December"; break;
                default: month = "Unknown"; break;
            }
            day = day.TrimStart(new Char[] { '0' }); // remove leading zero
            day = day.TrimEnd(new Char[] { '.' }); // remove last dot
            string date = year + "-" + month + "-" + day;
            appVer = (date).Trim();

            Text = "HovText Update (version "+ appVer +")";

            TopMost = true;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            CenterLabel(label3);
        }

        private void CenterLabel(Label label)
        {
            label.Location = new System.Drawing.Point(
                (this.ClientSize.Width - label.Width) / 2,
                (this.ClientSize.Height - label.Height) / 2
            );
            label.Anchor = AnchorStyles.None;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            secs--;
            label2.Text = secs.ToString();

            if(secs == 0)
            {
                timer1.Enabled = false;
                label1.Visible = false;
                label2.Visible = false;
                StartUpdate();
            }
        }

        private void StartUpdate()
        {

            // Get the argument passed - should be either "Development" or "Stable"
            string envType = "";
            if(startupArgs.Length > 0)
            {
                if (startupArgs[0].ToLower() == "stable")
                {
                    envType = "STABLE";
                } else
                {
                    envType = "DEVELOPMENT";
                }
            }

            // Get OS name
            string osVersion = (string)(from x in new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem").Get().Cast<ManagementObject>() select x.GetPropertyValue("Caption")).FirstOrDefault();
            bool closeNow = false;

            // Get the newest stable version
            string newestStableVersion = "";
            if (envType == "STABLE")
            {
                try
                {

                    WebClient webClient = new WebClient();
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
                    webClient.Headers.Add("user-agent", ("HovText Updater " + appVer).Trim());

                    // Prepare the POST data
                    var postData = new System.Collections.Specialized.NameValueCollection
                        {
                            { "osVersion", osVersion }
                        };

                    // Send the POST data to the server
                    byte[] responseBytes = webClient.UploadValues(hovtextPage + "/autoupdate/", postData);

                    // Convert the response bytes to a string
                    string newestStableVersionWithVersion = Encoding.UTF8.GetString(responseBytes);
                    string prefixToRemove = "Version: ";

                    newestStableVersion = newestStableVersionWithVersion;
                    if (newestStableVersionWithVersion.StartsWith(prefixToRemove))
                    {
                        newestStableVersion = newestStableVersionWithVersion.Substring(prefixToRemove.Length);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("HovText Updater error:\r\n\r\n" + ex.Message,
                        "HovText Updater ERROR",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    closeNow = true;
                }
            }

            if(!closeNow)
            {
            
                label3.Text = $"Fetching newest [{envType}] version from HovText home page.\n\nPlease wait ...";
                Refresh();

                // Set some variables
                string folderAndExe = envType == "DEVELOPMENT" ? "autoupdate/development/HovText.exe" : "download/" + newestStableVersion + "/HovText.exe";
                string updateExe = hovtextPage +"/"+ folderAndExe;
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string pathAndTempExe = Path.Combine(baseDirectory, "HovText.exe");

                // Delete the current "HovText.exe" as the application should have been closed by now
                if(File.Exists(baseDirectory + "HovText.exe"))
                {
                    try
                    {
                        File.Delete(baseDirectory + "HovText.exe");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("HovText Updater error:\r\n\r\n" + ex.Message,
                            "HovText Updater ERROR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        closeNow = true;
                    }
                }

                if(!closeNow)
                {
                    try
                    {
                        Thread.Sleep(1000);
                        WebClient webClient = new WebClient();
                        webClient.DownloadFile(updateExe, pathAndTempExe);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("HovText Updater error:\r\n\r\n" + ex.Message,
                            "HovText Updater ERROR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        closeNow = true;
                    }

                    // Start the new HovText executable
                    if(!closeNow)
                    {
                        Thread.Sleep(1000);
                        Process.Start(pathAndTempExe);
                    }
                }
            }

            // Terminate the main application
            Close();
        }
    }
}