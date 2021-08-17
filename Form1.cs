using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Threading.Tasks;

namespace Netsmart_Cache
{
    public partial class Netsmart_cache_application : Form
    {

        public Netsmart_cache_application()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            DeletOldBackups();
            cacheFolderReplace();

        }



        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }


        public async void cacheFolderReplace() 
        {

            //Variables for paths, backup counter and backup limit
            String date = DateString();
            String cachePath = @"C:\Program Files (x86)\Netsmart Homecare\Client\Cache";
            String backupCache = @"C:\Program Files (x86)\Netsmart Homecare\Client\CacheBackup_" + date;
            string[] dirs = Directory.GetDirectories(@"C:\Program Files (x86)\Netsmart Homecare\Client\", "CacheBackup_*");
            int backUpCount = 0;
            int maxBackups = 3;

            //Windows account name that the app is being run under
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString();

            //Computer Name
            string computerName = System.Environment.MachineName;

            //When the button is pressed the pop up message below appears. Change message string to edit the message and the title to change the title bar notice
            string message = "Are you sure you want to exit Netsmart and clear the cache? Press yes to continue.";
            string title = "Warning";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                //Close Netsmart and pause the program to give Netsmart time to stop
                CloseNetsmart();
                System.Threading.Thread.Sleep(5000);


                // Try catch attempts to rename and add cache folder. If errors are thrown it will assume Netsmart is open(read/write error) and instead ask the person to close out of Netsmart.
                try
                {

                    foreach (string dir in dirs)
                    {
                        backUpCount++;
                    }

                    //This looks to see if a backup exists for the day. If it does then it will not proceed and instead ask them to contact the help desk. I assume that if they are trying to run the app
                    //multiple times in one day that there is something else going on that they should call for
                    if (Directory.Exists(backupCache))
                    {

                        string errMessage = "A backup already exists for today's Netsmart data. Follow the how to guide for directions on syncing Netsmart back into field mode. If you are still experiencing issues after following those steps, please contact the helpdesk at support@caredimensions.org or on extension 4357";
                        string errTitle = "Cannot reset data";
                        MessageBox.Show(errMessage, errTitle);
                        //this.Close();

                    } 
                    else if (backUpCount >= maxBackups)
                    {
                       string errMessage = "There are too many backup files present. To ensure no loss of patient data, please contact the helpdesk at support@caredimensions.org or on extension 4357";
                       string errTitle = "Cannot reset data";
                       MessageBox.Show(errMessage, errTitle);
                       //this.Close();
                    } 
                    else {

                        //If the program runs this far then it will rename the cache with the date and then create a new cache folder
                        System.IO.Directory.Move(@"C:\Program Files (x86)\Netsmart Homecare\Client\Cache", @"C:\Program Files (x86)\Netsmart Homecare\Client\CacheBackup_" + date);
                        DirectoryInfo di = Directory.CreateDirectory(cachePath);


                        string subject = userName + " has performed a cache clear";
                        string body = "Notice: The following user has succesfully cleared their Netsmart cache data." + "\nUser: " + userName + "\nComputer name: " + computerName + "\nDate: " + date;

                        await SendEmailTicket(subject, body);

                        string sMessage = "Your cached data has been reset! If you work in field mode, please follow the user guide to sync back into field mode. If this has not fixed your issue, please contact the helpdesk at support@caredimensions.org or on extension 4357";
                        string sTitle = "Operation Successful";
                        MessageBox.Show(sMessage, sTitle);
                        //this.Close();

                    }
                }
                catch
                {
                    //Catches any error and tells the user to close netsmart(assuming read/write error).       
                    string eMessage = "We could not finish the process. Please exit out of Netsmart and then try this tool again. If this error persists, contact the helpdesk at support@caredimensions.org or on extension 4357";
                    string eTitle = "Error";
                    MessageBox.Show(eMessage, eTitle);
                    //this.Close();
                }

            }
            else
            {
                
                this.Close();
            }
        }

        //This gets the current date in the format specified and converts it to a string so it can be used to rename the cache folder.
        public String DateString() 
        {
            DateTime dt = DateTime.Now;
            String date = dt.ToString("MM-dd-yy", DateTimeFormatInfo.InvariantInfo);
            return date;
        }

        public String[] DatesToCompare()
        {

            //This function creates an array of dates as strings that can be used to filter old backup folders

            //Creates an array for weeks worth of days to compare. maxDays is set to 0 because we first set todays date first and then add the rest of the days. 
            //Any more than 6 will create an error since we are only initializing an array with 7 places
            string[] dates = new string[7];
            int maxDays = 6;

            DateTime dt = DateTime.Now;
            dates[0] = dt.ToString("MM-dd-yy", DateTimeFormatInfo.InvariantInfo);

            for (int d = 1; d < maxDays; d++)
            {
                DateTime tempdt = DateTime.Now.Date.AddDays(-d);
                dates[d] = tempdt.ToString("MM-dd-yy", DateTimeFormatInfo.InvariantInfo);
            }
            
            return dates;
        }

        public void DeletOldBackups()
        {
            //String arrays that hold the list of found backup folders and our list of days 
            string[] dirs = Directory.GetDirectories(@"C:\Program Files (x86)\Netsmart Homecare\Client\", "CacheBackup_*");
            string[] days = DatesToCompare();

            //Loop through all the backup folders that are found
            foreach (string dir in dirs)
            {
                //Assuming a backup folder is too old
                bool tooOld = true;

                //Loop condition that goes through our list of recent days and compares if the folder name matches
                //If the folder is a match then tooOld is false
                for (int x = 0; x < days.Length; x++)
                {
                    String backupCache = @"C:\Program Files (x86)\Netsmart Homecare\Client\CacheBackup_" + days[x];
                    if (backupCache == dir)
                    {
                        tooOld = false;
                    }
                }

                //If tooOld is still true after checking the dates then the directory is deleted
                if (tooOld)
                {
                    string[] filePaths = Directory.GetFiles(dir);
                    foreach (string filePath in filePaths)
                    {
                        File.Delete(filePath);
                    }
                    Directory.Delete(dir);
                }
            }
        }


        public void CloseNetsmart()
        {
            //Close all Netsmart(MHC) processes
            var processArray = Process.GetProcesses();
            var process = processArray.FirstOrDefault(p => p.ProcessName == "MHC");
            process?.Kill();
        }


        public async Task SendEmailTicket(string messageSubject, string messageBody)
        {
            //Certificate workaround
            ServicePointManager.ServerCertificateValidationCallback =
            delegate (object sender, X509Certificate certificate, X509Chain chain,
            SslPolicyErrors sslPolicyErrors) { return true; };

            //Smtp email server and account settings
            var smtpClient = new SmtpClient("exchange server")
            {
                Port = 25,
                Credentials = new NetworkCredential("emailaddress", "password"),
                EnableSsl = true,
            };

            //Email parameters
            string body = messageBody;
            string subject = messageSubject;
            string sendFrom = "hdesk@caredimensions.org";
            string sendTo = "hdesk@caredimensions.org";

            //Sends the email
            await smtpClient.SendMailAsync(sendFrom, sendTo, subject, body);
        }


        //TODO - fix link to display sync directions pdf
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"P:\Liberty Public\IT\Tips and Tricks\How to sync into field mode.pdf");
        }
    }
}
