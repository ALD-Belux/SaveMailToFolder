using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using System.Text.RegularExpressions;
using Serilog;

namespace SaveMailToFolderLibrary
{
    public class RetrieveMails
    {

        private string mailboxSMTP;
        private ExchangeService service;

        public RetrieveMails(string mailboxSMTP, ExchangeVersion requestedServerVersion = ExchangeVersion.Exchange2010_SP2)
        {
            try
            {
                Log.Debug("ExchangeService: Creating");
                ExchangeService service = new ExchangeService(requestedServerVersion);
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(mailboxSMTP);
                this.service = service;
                this.mailboxSMTP = mailboxSMTP;
                Log.Debug("ExchangeService: Ready");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed while trying to create ExchangeService");
                throw;
            }
            
        }

        public string sendTest(string recipient)
        {

            try
            {
                EmailMessage message = new EmailMessage(service);
                message.Subject = "[C#] Test Message";
                message.Body = "Hello EWs World!";
                message.ToRecipients.Add(recipient);
                message.From = mailboxSMTP;
                message.SendAndSaveCopy();

                return "Message sent";

            }
            catch (Exception ex)
            {

                return "Error: " + ex;
            }
            

        }

        public void SaveInboxToEml(string path, bool delete = false, bool saveToEML = true)
        {
            Log.Debug("Entering SaveInboxToEml with path = {savePath} and delete = {delete}", path, delete);
            int offset = 0;
            int pageSize = 50;
            bool more = true;

            Mailbox mb = new Mailbox(mailboxSMTP);
            FolderId fIdInbox = new FolderId(WellKnownFolderName.Inbox, mb);

            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);
            view.PropertySet = PropertySet.IdOnly;

            FindItemsResults<Item> findResults;
            List<ItemId> allResultsID = new List<ItemId>();
            List<EmailMessage> emails = new List<EmailMessage>();

            try
            {
                Log.Debug("Start mails query");
                while (more)
                {
                    findResults = service.FindItems(fIdInbox, view);
                    Log.Debug("Pass {FindItemIteration} - findResults: {findResults}", view.Offset / pageSize, findResults.Count());
                    foreach (var item in findResults.Items)
                    {
                        emails.Add((EmailMessage)item);
                    }
                    more = findResults.MoreAvailable;
                    if (more)
                    {
                        view.Offset += pageSize;
                    }
                }                

                if (emails.Count > 0)
                {
                    Log.Debug("{nbEmails} founds. Get properties.", emails.Count);
                    
                    //PropertySet properties = (BasePropertySet.FirstClassProperties); //A PropertySet with the explicit properties you want goes here
                    PropertySet properties = new PropertySet(ItemSchema.Subject, ItemSchema.MimeContent);
                    service.LoadPropertiesForItems(emails, properties);

                    Log.Debug("Save to eml enabled: {saveToEML}", saveToEML);

                    foreach (EmailMessage msg in emails)
                    {                        
                        MimeContent mimcon = msg.MimeContent;
                        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string filename = MakeValidFileName(CleanSubject(msg.Subject));

                        if (path.Length + filename.Length > 230)
                        {
                            filename = filename.Substring(0, ( 230 - path.Length));
                        }

                        string emlFile = string.Format("{0}\\{1}-{2}.eml", path, filename, timestamp);

                        Log.Information("Message with subject: {Subject} will be saved to {file}", msg.Subject, emlFile);

                        try
                        {
                            if(saveToEML)
                            {
                                FileStream fStream = new FileStream(emlFile, FileMode.Create);
                                fStream.Write(mimcon.Content, 0, mimcon.Content.Length);
                                fStream.Close();
                            }

                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Failed to save eml file.");
                            throw;
                        }

                        allResultsID.Add(msg.Id);
                    }

                    Log.Debug("Delete mail? {DeleteMail}", delete);
                    if(delete)
                    {
                        Log.Information("Start Bulk Delete of {nbMailsToDelete} mails", allResultsID.Count);
                        service.DeleteItems(allResultsID, DeleteMode.SoftDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                    }
                }                
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to retrieve mails.");
                throw;
            }
        }

        private static string MakeValidFileName(string name)
        {
            var invalidFileNameChars = new string(Path.GetInvalidFileNameChars());
            string invalidChars = Regex.Escape(invalidFileNameChars);
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return Regex.Replace(name, invalidRegStr, "_");
        }

        private static string CleanSubject(string subject)
        {
            switch (subject.ToLower().Substring(0,3))
            {
                case "re:":
                case "fw:":
                    return subject.Substring(3, subject.Length - 3).Trim();
                default:
                    return subject.Trim();
            }
        }
    }
}
