using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using SaveMailToFolderLibrary;
using Serilog;


namespace SaveMailToFolderService
{
    public partial class SaveMailSVC : ServiceBase
    {

        string mailboxSMTP = Properties.Settings.Default.MailboxSMTP;
        string testMail = Properties.Settings.Default.RecipientTest;
        bool deleteWhenSaved = Properties.Settings.Default.DeleteWhenSaved;
        Microsoft.Exchange.WebServices.Data.ExchangeVersion requestedServerVersion = Properties.Settings.Default.ExchangeVersion;
        string savePath = Properties.Settings.Default.SavePath;
        bool saveToEML = Properties.Settings.Default.SaveToEML;


        bool methodRunning = false;
        Timer timer = new Timer(Properties.Settings.Default.ProcessIntervalMs);
        RetrieveMails rtMail;

        public SaveMailSVC()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            bool run = true;
            while(run)
            {
                run = !(Start());
            }
        }

        public bool Start()
        {
            try
            {
                rtMail = new RetrieveMails(mailboxSMTP, requestedServerVersion);
            }
            catch (Exception)
            {
                Log.Error("An error occured, please review previous logs");
                return false;
            }
           
            timer.Elapsed += Timer_Elapsed;
            timer.Enabled = true;
            return true;
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            methodRunning = true;
            try
            {                
                rtMail.SaveInboxToEml(savePath, deleteWhenSaved, saveToEML);                
            }
            catch (Exception)
            {
                Log.Error("An error occured, please review previous logs");
            }
            methodRunning = false;
        }

        protected override void OnStop()
        {
            timer.Enabled = false;

            while (methodRunning)
            {
                Log.Information("Waiting method to finish");
                System.Threading.Thread.Sleep(1000);
            }
            Log.Information("Stopping Service");
        }
    }
}


