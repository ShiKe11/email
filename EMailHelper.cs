using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Net.Mail;
using System.Net;
using System.Configuration;
using System.Xml;

namespace AutomationHelper
{
    public class EMailHelper
    {


        public static void SendMail(string title, string target, string text, bool isHtml)
        {
            string subject = string.Format("{0}:{1}", title, CommonHelper.GetNowString());
            //            string from = "qinguo@microsoft.com";
            string from = System.Environment.UserName + "@microsoft.com";
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(from);
            List<string> tos = ParseToTarget(target);
            foreach (string to in tos)
            {
                msg.To.Add(to);
            }
            msg.Subject = subject;
            msg.Body = text;
            msg.IsBodyHtml = isHtml;
            SmtpClient client = new SmtpClient("smtphost.redmond.corp.microsoft.com");
            client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
            client.Send(msg);
        }

        public static void SendMail(string title, string target, string text)
        {
            SendMail(title, target, text, false);
        }

        private static List<string> ParseToTarget(string target)
        {
            string[] nodes = target.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> deduped = new List<string>();
            foreach (string node in nodes)
            {
                if (!deduped.Contains(node))
                {
                    deduped.Add(node);
                }
            }
            return deduped;
        }

    }
}
