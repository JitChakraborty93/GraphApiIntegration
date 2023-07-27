//using Azure.Identity;
//using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace GraphApiMailService
{
    public partial class Service1 : ServiceBase
    {
        SqlConnection sqlConnection = new SqlConnection(ConfigurationSettings.AppSettings["DBConnection"].ToString());
        SqlDataAdapter sqlDataAdapter;
        SqlCommand sqlCommand;
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
            string TopFolderCount = ConfigurationSettings.AppSettings["TopFolderCount"].ToString();
            string MailboxReadfolder = ConfigurationSettings.AppSettings["MailboxReadfolder"].ToString();
            string MailboxArchivefolder = ConfigurationSettings.AppSettings["MailboxArchivefolder"].ToString();
            DataTable dtmailConfig = new DataTable();
            dtmailConfig = Fetusername(username);
            string password = dtmailConfig.Rows[0]["Password"].ToString();
            try
            {
                //var credential= new ClientSecretCredential(tenantId, 
                //    clientId,
                //    ClientSecretId, 
                //    new TokenCredentialOptions { AuthorityHost=AzureAuthorityHosts.AzurePublicCloud});
                //GraphServiceClient  graphServiceClient = new GraphServiceClient(credential);
                //var allMailFolder = graphServiceClient.Users[username].MailFolders.Request().GetAsync().Result;
                //foreach (MailFolder item in allMailFolder)
                //{

                //}
            }
            catch (Exception)
            {

                
            }

        }

        protected override void OnStop()
        {
        }

        #region Database 
        public DataTable Fetusername(string username)
        {
            DataTable dt = new DataTable();
            sqlDataAdapter = new SqlDataAdapter();
            sqlCommand = new SqlCommand("SP_Mail_Configuration",sqlConnection);
            sqlCommand.CommandType = CommandType.StoredProcedure;
            SqlParameter paramusername = new SqlParameter("@username", username);
            sqlCommand.Parameters.Add(paramusername);
            sqlDataAdapter.SelectCommand = sqlCommand;
            sqlConnection.Open();
            sqlDataAdapter.Fill(dt);
            sqlConnection.Close();
            return dt;

        }


        #endregion

    }
}
