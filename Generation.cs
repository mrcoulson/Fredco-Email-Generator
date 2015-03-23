using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Email4
{
    /// <summary>
    /// Methods for generating an email. By Jeremy Coulson.
    /// </summary>
    public class Generation
    {
        /// <summary>
        /// Sends an email with more options.
        /// </summary>
        /// <param name="strFrom">The from field. Required.</param>
        /// <param name="strTo">Recipients. Separate with comma. At least one is required.</param>
        /// <param name="strCC">CC recipients. Separate with comma. Optional; use null or Nothing if blank.</param>
        /// <param name="strAttachments">Attachments. Separate with comma. Optional; use null or Nothing if blank.</param>
        /// <param name="strSubject">The subject of the message. Required.</param>
        /// <param name="strBody">The body of the message. Required.</param>
        /// <param name="strReplyTo">Reply recipient. Only list one. On IIS 7, this requires integreated pipeline mode. Optional; use null or Nothing if blank.</param>
        /// <param name="boolBodyHtml">Is the body HTML or not?  Required.</param>
        /// <param name="strMailServer">The SMTP client to send this message. Required.</param>
        /// <param name="strUser">The SMTP user. Required.</param>
        /// <param name="strPass">The SMPT password. Required.</param>
        /// <param name="boolSsl">True to connect with SSL. False to connect without SSL. Required.</param>
        /// <returns>Retuns success or exception.</returns>
        public string SendEmail(string strFrom, string strTo, string strCC, string strAttachments, string strSubject, string strBody, string strReplyTo, bool boolBodyHtml, string strMailServer, string strUser, string strPass, bool boolSsl)
        {
            // String to hold result message.
            string strEmailSuccess = String.Empty;

            try
            {
                using (System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage())
                {
                    msg.IsBodyHtml = boolBodyHtml;
                    msg.From = new System.Net.Mail.MailAddress(strFrom);
                    // Split the To string into an array and add each address.
                    string[] strToArray = strTo.Split(',');
                    foreach (string strToRecipient in strToArray)
                    {
                        msg.To.Add(new System.Net.Mail.MailAddress(strToRecipient));
                    }
                    if ((strCC != null) && (strCC != String.Empty) && (strCC.Length > 0))
                    {
                        // Split the CC string into an array and add each address.
                        string[] strCCArray = strCC.Split(',');
                        foreach (string strCCRecipient in strCCArray)
                        {
                            msg.CC.Add(new System.Net.Mail.MailAddress(strCCRecipient));
                        }
                    }
                    msg.Subject = strSubject;
                    msg.Body = strBody;
                    if ((strReplyTo != null) && (strReplyTo != String.Empty) && (strReplyTo.Length > 0))
                    {
                        msg.ReplyTo = new System.Net.Mail.MailAddress(strReplyTo);
                    }
                    if ((strAttachments != null) && (strAttachments != String.Empty) && (strAttachments.Length > 0))
                    {
                        // Split the attachments string into an array and add each attachment.
                        string[] strAttachmentArray = strAttachments.Split(',');
                        foreach (string strAttachment in strAttachmentArray)
                        {
                            System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(strAttachment);
                            msg.Attachments.Add(attachment);
                        }
                    }
                    using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(strMailServer))
                    {
                        System.Net.NetworkCredential user = new System.Net.NetworkCredential(strUser, strPass);
                        client.Credentials = user;
                        client.EnableSsl = boolSsl;
                        client.Send(msg);
                    }
                }
                // Success!  Say so.
                strEmailSuccess = "Complex overload sent at " + System.DateTime.Now.ToString() + "!";
            }
            catch (Exception ex)
            {
                // Problem!  Say so.
                strEmailSuccess = ex.ToString();
            }

            return strEmailSuccess;
        }

        /// <summary>
        /// Sends an email with fewer options. SMTP server info comes from Web.config.
        /// </summary>
        /// <param name="strFrom">The from field. Required.</param>
        /// <param name="strTo">Recipients. Separate with comma. Required.</param>
        /// <param name="strCC">CC recipients. Separate with comma. Optional; use null or Nothing if blank.</param>
        /// <param name="strAttachments">Attachments. Separate with comma. Optional; use null or Nothing if blank.</param>
        /// <param name="strSubject">The subject of the message. Required.</param>
        /// <param name="strBody">The body of the message. Required.</param>
        /// <returns>Retuns success or exception.</returns>
        public string SendEmail(string strFrom, string strTo, string strCC, string strAttachments, string strSubject, string strBody)
        {
            // Create string to hold success message.
            string strEmailSuccess = String.Empty;

            try
            {
                using (System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage())
                {

                    msg.IsBodyHtml = false;
                    msg.From = new System.Net.Mail.MailAddress(strFrom);
                    // Split the To string into an array and add each address.
                    string[] strToArray = strTo.Split(',');
                    foreach (string strToRecipient in strToArray)
                    {
                        msg.To.Add(new System.Net.Mail.MailAddress(strToRecipient));
                    }
                    if ((strCC != null) && (strCC != String.Empty) && (strCC.Length > 0))
                    {
                        // Split the CC string into an array and add each address.
                        string[] strCCArray = strCC.Split(',');
                        foreach (string strCCRecipient in strCCArray)
                        {
                            msg.CC.Add(new System.Net.Mail.MailAddress(strCCRecipient));
                        }
                    }
                    msg.Subject = strSubject;
                    msg.Body = strBody;
                    if ((strAttachments != null) && (strAttachments != String.Empty) && (strAttachments.Length > 0))
                    {
                        // Split the attachments string into an array and add each attachment.
                        string[] strAttachmentArray = strAttachments.Split(',');
                        foreach (string strAttachment in strAttachmentArray)
                        {
                            System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(strAttachment);
                            msg.Attachments.Add(attachment);
                        }
                    }
                    using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(System.Web.Configuration.WebConfigurationManager.AppSettings["GlobalSMTP"].ToString()))
                    {
                        System.Net.NetworkCredential user = new System.Net.NetworkCredential(System.Web.Configuration.WebConfigurationManager.AppSettings["GlobalSMTPUser"].ToString(), System.Web.Configuration.WebConfigurationManager.AppSettings["GlobalSMTPPass"].ToString());
                        client.Credentials = user;
                        client.EnableSsl = true;
                        client.Send(msg);
                    }
                    strEmailSuccess = "Simple overload sent at " + System.DateTime.Now.ToString() + "!";
                }
            }
            catch (Exception ex)
            {
                strEmailSuccess = ex.ToString();
            }

            return strEmailSuccess;
        }
    }
}
