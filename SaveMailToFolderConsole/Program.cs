using SaveMailToFolderLibrary;
using Serilog;
using System;

namespace SaveMailTo_FolderConsole
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string mailboxSMTP = Properties.Settings.Default.MailboxSMTP;
            string testMail = Properties.Settings.Default.RecipientTest;
            bool deleteWhenSaved = Properties.Settings.Default.DeleteWhenSaved;
            Microsoft.Exchange.WebServices.Data.ExchangeVersion requestedServerVersion = Properties.Settings.Default.ExchangeVersion;
            string savePath = Properties.Settings.Default.SavePath;
            bool saveToEML = Properties.Settings.Default.SaveToEML;

            Log.Logger = new LoggerConfiguration()
                            .ReadFrom.AppSettings()
                            .CreateLogger();

            Log.Debug("Hello Serilog!");
            Log.Information("Starting Process");

            try
            {
                RetrieveMails rtMail = new RetrieveMails(mailboxSMTP, requestedServerVersion);
                rtMail.SaveInboxToEml(savePath, deleteWhenSaved, saveToEML);
            }
            catch (Exception)
            {
                Log.Error("An error occured, please review previous logs");
            }

            Log.Information("Success");
            Console.ReadLine();
        }
    }
}