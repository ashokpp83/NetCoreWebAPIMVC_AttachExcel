using NLog;
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace WebAPIMVC_AttachExcel
{
    /// <summary>
    /// EmailService
    /// </summary>
    public class EmailService : IEmailService
    {
        private static readonly Logger logMe = LogManager.GetCurrentClassLogger();
        /// <summary>
        /// support_EmailList
        /// </summary>
        public string support_EmailList = "";

        /// <summary>
        /// smtp_Server
        /// </summary>
        public string smtp_Server = "";

        /// <summary>
        /// smtp_Port
        /// </summary>
        public int smtp_Port = 0;


        public EmailService(Dictionary<string, string> AppSettings)
        {
            support_EmailList = AppSettings["Support_EmailList"];
            smtp_Server = AppSettings["SMTP_Server"];
            smtp_Port = Convert.ToInt32(AppSettings["SMTP_Port"]);
        }
        /// <summary>
        /// SendEmailNotification
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="msgBody"></param>
        /// <param name="logMe"></param>
        public void SendEmailNotification(string subject, string msgBody)
        {
            try
            {
                var _smtpMailer = new SmtpClient(smtp_Server);
                _smtpMailer.Port = smtp_Port;
                _smtpMailer.UseDefaultCredentials = true;

                MailMessage MailMessage = new MailMessage();
                MailMessage.From = new MailAddress(support_EmailList);

                MailMessage.To.Add(support_EmailList);
                //if (!string.IsNullOrEmpty(ccTo))
                //{
                //    MailMessage.CC.Add(ccTo);
                //}

                MailMessage.IsBodyHtml = false;
                MailMessage.Subject = subject;
                MailMessage.Body = msgBody;


                _smtpMailer.Send(MailMessage);

            }
            catch (Exception ex)
            {
                logMe.Error("Exception occured at SendEmailNotification: " + ex.Message);
            }
        }

    }
}
