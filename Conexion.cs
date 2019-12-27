

using System;
using System.Data;
using System.Data.SqlClient;
using Entidades;
using System.Configuration;
using System.Net.Mail;
using System.Text;
using System.Net;

namespace Data
{
    public class Conexion : IDisposable
    {

        private SqlConnection _sqlconnnection = null;

        virtual protected string ConnectionString()
        {
            return System.Configuration.ConfigurationManager.AppSettings["appConnectionString"];
        }

        //public static string GetConnectionString()
        //{
        //    return System.Configuration.ConfigurationManager.AppSettings["appConnectionString"];
        //}


        protected SqlConnection Conn
        {
            get
            {
                return _sqlconnnection;
            }
        }

        protected void open_Conn()
        {
            if (_sqlconnnection == null)
            {
                _sqlconnnection = new SqlConnection(ConnectionString());
                _sqlconnnection.Open();
            }
            int i = 0;
            while (_sqlconnnection.State != ConnectionState.Open)
            {
                if (i < 5)
                {
                    try
                    {
                        _sqlconnnection.Open();
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(10);
                    }
                }
            }
        }

        protected void Close_Conn()
        {
            try
            {
                _sqlconnnection.Close();
                _sqlconnnection.Dispose();
                _sqlconnnection = null;
            }
            catch
            { }
        }


        /// <summary>
        /// Open Connection and Get the Paramenters of the SP
        /// </summary>
        /// <param name="strSPName"></param>
        /// <returns></returns>
        protected SqlCommand GetSqlCommandInstance(string strSPName)
        {
            open_Conn();
            System.Data.SqlClient.SqlCommand sqlcommand = new System.Data.SqlClient.SqlCommand(strSPName, Conn);
            sqlcommand.CommandType = CommandType.StoredProcedure;
            sqlcommand.CommandTimeout = 1800;
            return sqlcommand;
        }

        protected object ScalarSQLCommand(System.Data.SqlClient.SqlCommand sqlcommand)
        {
            object ret = sqlcommand.ExecuteScalar();
            DisposeCommand(sqlcommand);
            return ret;
        }

        protected int ReturnValueSQLCommand(System.Data.SqlClient.SqlCommand sqlcommand)
        {
            int returnValue = sqlcommand.ExecuteNonQuery();

            DisposeCommand(sqlcommand);

            return returnValue;
        }

        protected void DisposeCommand(System.Data.SqlClient.SqlCommand sqlcommand)
        {
            sqlcommand.Dispose();
            Close_Conn();
        }

        public void Dispose()
        {
            //Dispose(true);
            //GC.SuppressFinalize(this);
        }


        public static Email GetConfigurationEmail()
        {
            Email email = new Email();
            int port = 0;
            int.TryParse(ConfigurationManager.AppSettings["MailPort"], out port);

            email.EmailServerBe = new EmailServerBE()
            {
                Host = ConfigurationManager.AppSettings["MailSmtpServer"],
                User = ConfigurationManager.AppSettings["MailUserName"],
                Password = ConfigurationManager.AppSettings["MailPassword"],
                From = ConfigurationManager.AppSettings["MailFrom"],
                Port = port
            };

            return email;
        }

        public static bool SendEmail(Email mail)
        {
            //try
            //{
                MailMessage message = new MailMessage();
                string str = "<html><body><font face='Arial' size='-1'>";
                str = (str + mail.EmalBody + "</font></body></html>");
                message.BodyEncoding = Encoding.UTF8;
                message.Body = str;
                message.IsBodyHtml = true;
                message.From = new MailAddress(mail.EmailServerBe.From);

                string[] strToSplit = mail.EmailTo.Split(new string[] { ";", "," }
                    , StringSplitOptions.RemoveEmptyEntries);

                foreach (string strEmail in strToSplit)
                {
                    if (!string.IsNullOrEmpty(strEmail.Trim()))
                    {
                        message.To.Add(strEmail);
                    }
                }

                if (!string.IsNullOrEmpty(mail.EmailCC))
                {
                    string[] strBCCSplit = mail.EmailCC.Split(new string[] { ";", "," }
                        , StringSplitOptions.RemoveEmptyEntries);
                    foreach (string strEmail in strBCCSplit)
                    {
                        if (!string.IsNullOrEmpty(strEmail.Trim()))
                        {
                            message.Bcc.Add(strEmail);
                        }
                    }
                }
                
                message.Subject = mail.EmailSubject;
                SmtpClient client = new SmtpClient();
                client.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["IsSSL"]);        
                client.Host = mail.EmailServerBe.Host;
                client.Port = mail.EmailServerBe.Port;
                client.Credentials = new NetworkCredential(mail.EmailServerBe.User, mail.EmailServerBe.Password);
                client.Send(message);
                return (true);
            //}
            //catch
            //{

            //    return false;
            //}

        }




    }
}
