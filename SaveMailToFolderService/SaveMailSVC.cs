using SaveMailToFolderLibrary;
using Serilog;
using System;
using System.ServiceProcess;
using System.Timers;

namespace SaveMailToFolderService
{
    public partial class SaveMailSVC : ServiceBase
    {
        private string mailboxSMTP = Properties.Settings.Default.MailboxSMTP;
        private string testMail = Properties.Settings.Default.RecipientTest;
        private bool deleteWhenSaved = Properties.Settings.Default.DeleteWhenSaved;
        private Microsoft.Exchange.WebServices.Data.ExchangeVersion requestedServerVersion = Properties.Settings.Default.ExchangeVersion;
        private string savePath = Properties.Settings.Default.SavePath;
        private bool saveToEML = Properties.Settings.Default.SaveToEML;

        private bool methodRunning = false;
        private Timer timer = new Timer(Properties.Settings.Default.ProcessIntervalMs);
        private RetrieveMails rtMail;

        public SaveMailSVC()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            bool run = true;
            while (run)
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