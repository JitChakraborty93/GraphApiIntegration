//using Azure.Identity;
//using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace GraphApiMailService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            MailRead();    

        }

        public void MailRead()
        {
            string username = ConfigurationSettings.AppSettings["UserName"].ToString();
            string tenantId = ConfigurationSettings.AppSettings["TenantId"].ToString();
            string clientId = ConfigurationSettings.AppSettings["ClientId"].ToString();
            string ClientSecretId = ConfigurationSettings.AppSettings["ClientSecretId"].ToString();
            int TopFolderCount = Convert.ToInt32(ConfigurationSettings.AppSettings["TopFolderCount"].ToString());
            int TopMailCount = Convert.ToInt32(ConfigurationSettings.AppSettings["TopMailCount"].ToString());
            string MailboxReadfolder = ConfigurationSettings.AppSettings["MailboxReadfolder"].ToString();
            string MailboxArchivefolder = ConfigurationSettings.AppSettings["MailboxArchivefolder"].ToString();
            string LocalFoldarPathAttachment = ConfigurationSettings.AppSettings["LocalFoldarPathAttachment"].ToString();
            string LocalFoldarPath = ConfigurationSettings.AppSettings["LocalFoldarPath"].ToString();
            try
            {
                var credential = new ClientSecretCredential(tenantId,
                    clientId,
                    ClientSecretId,
                    new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud });
                GraphServiceClient graphServiceClient = new GraphServiceClient(credential);
                var allMailFolder = graphServiceClient.Users[username].MailFolders.Request().GetAsync().Result;
                foreach (MailFolder folder in allMailFolder)
                {
                    if (folder.DisplayName == "Inbox")//in My case my folder from where my service will start read under the inbox folder in your case if folder is in rootfolder then check with same
                    {
                        var childFolders = graphServiceClient.Users[username].MailFolders["Inbox"].ChildFolders.Request().Top(TopFolderCount).GetAsync().Result;
                        //TopFolderCount will set from app.config cause by default pick top 10 folder to fix this we need max folder count
                        foreach (MailFolder subfolder in childFolders)
                        {
                            if (subfolder.DisplayName == MailboxReadfolder)
                            {
                                var subId = subfolder.Id.ToString();//you have to get the folder id to process further
                                var fmessage = graphServiceClient.Users[username].MailFolders["Inbox"].ChildFolders[subId].Messages.Request().Top(TopMailCount).GetAsync().Result;
                                //TopFolderCount will set from app.config cause by default pick top 10 mails to fix this we need max mail count
                                string mailid = string.Empty, cc = string.Empty, to = string.Empty, body = string.Empty, subject = string.Empty, subjectWithDate = string.Empty
                                , fPath = string.Empty, from = string.Empty;
                                foreach (var folderMessage in fmessage)
                                {
                                    fPath = LocalFoldarPathAttachment;
                                    Message message = (Message)folderMessage;
                                    from = folderMessage.From.EmailAddress.Address.ToString();
                                    if (message.ToRecipients != null)
                                    {
                                        foreach (Recipient item in message.ToRecipients)
                                        {
                                            to = item.EmailAddress.Address.ToString() + ";";
                                        }
                                    }
                                    if (message.CcRecipients != null)
                                    {
                                        foreach (Recipient item in message.CcRecipients)
                                        {
                                            cc = item.EmailAddress.Address.ToString() + ";";
                                        }
                                    }
                                    var msgid = message.Id.ToString();
                                    var attach = graphServiceClient.Users[username].MailFolders["Inbox"].ChildFolders[subId].Messages[msgid].Attachments.Request().GetAsync().Result;
                                    #region Attachment
                                    foreach (var attachment in attach)
                                    {
                                        if (attachment is Microsoft.Graph.FileAttachment)
                                        {
                                            Microsoft.Graph.FileAttachment fileat = attachment as Microsoft.Graph.FileAttachment;
                                            byte[] filecontent = fileat.ContentBytes;
                                            string stemp = DateTime.Now.ToString("yyyyMMddHHmmss") + fileat.Name;
                                            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                                            {
                                                stemp = stemp.Replace(c, '_');
                                            }
                                            string savedAttachment = "";
                                            savedAttachment += stemp + ";";
                                            System.IO.File.WriteAllBytes(LocalFoldarPathAttachment + stemp, filecontent);
                                            if (savedAttachment == "")
                                            {
                                                savedAttachment = fileat.Name;
                                            }
                                            else
                                            {
                                                savedAttachment += ";" + fileat.Name;
                                            }
                                        }
                                        else
                                        {
                                            Microsoft.Graph.ItemAttachment fileat = attachment as Microsoft.Graph.ItemAttachment;
                                            string stemp = DateTime.Now.ToString("yyyyMMddHHmmss") + fileat.Name;
                                            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                                            {
                                                stemp = stemp.Replace(c, '_');
                                            }
                                            string savedAttachment = "";
                                            savedAttachment += stemp + fileat.Name.Split('.')[1].ToString();//to get the attachment extention
                                            fileat = message.Attachments.OfType<Microsoft.Graph.ItemAttachment>().SingleOrDefault(c => c.Name == savedAttachment);
                                            var bas64 = fileat.ContentType;
                                            Microsoft.Graph.MimeContent mcattach = new Microsoft.Graph.MimeContent();
                                            byte[] bytes = mcattach.Value;
                                            FileStream fileStream = new FileStream(LocalFoldarPathAttachment + stemp + fileat.Name.Split('.')[1].ToString(), FileMode.Create);
                                            fileStream.Write(bytes, 0, bytes.Length);
                                            fileStream.Close();
                                        }
                                        #endregion
                                        #region StoreMail
                                        if (folderMessage.Subject != null)
                                        {
                                            subject = folderMessage.Subject;
                                            subjectWithDate = DateTime.Now.ToString("yyyyMMddHHmmss") + subject;
                                            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                                            {
                                                subject = subject.Replace(c, '_');
                                                subjectWithDate = subjectWithDate.Replace(c, '_');
                                            }
                                            mailid = message.InternetMessageId;
                                            body = folderMessage.Body.Content == null ? "" : folderMessage.Body.Content.ToString();
                                            string dpath;
                                            dpath = LocalFoldarPath.Replace(LocalFoldarPath.Substring(0, LocalFoldarPath.Length), "");
                                            string day = message.ReceivedDateTime.Value.Date.ToString().Replace("/", "-").Substring(0, 10);
                                            string hh = message.ReceivedDateTime.Value.Hour.ToString();
                                            string mm = message.ReceivedDateTime.Value.Minute.ToString();
                                            string ss = message.ReceivedDateTime.Value.Second.ToString();
                                            if (mm.Length == 1)
                                            {
                                                mm = "0" + mm;
                                            }
                                            DateTime dt1 = Convert.ToDateTime((day + " " + hh + ":" + mm + ":" + ss).ToString());
                                            //StoremailinurDb Pass the required parameter
                                            //if store get successful you can move the mail from reading folder to archive folder

                                            var mimecontent = graphServiceClient.Users[username].Messages[msgid].Content.Request().GetAsync().GetAwaiter().GetResult();
                                            using (var filsteram = System.IO.File.Create(LocalFoldarPath + subjectWithDate + ".eml"))
                                            {
                                                mimecontent.Seek(0, SeekOrigin.Begin);
                                                mimecontent.CopyTo(filsteram);
                                            }
                                            ///once mail file got saved in local and then move the mail in archive folder
                                            string sarchive = MailboxArchivefolder;
                                            foreach (MailFolder fldr in childFolders)
                                            {
                                                if (fldr.DisplayName.Equals(MailboxArchivefolder))
                                                {
                                                    var fid = fldr.Id.ToString();
                                                    var msgmove = graphServiceClient.Users[username].Messages[msgid].Move(fid).Request().PostAsync().Result;
                                                    break;
                                                }
                                            }
                                            to = ""; cc = "";
                                            #endregion
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {


            }

        }

        protected override void OnStop()
        {
        }

    }
}
