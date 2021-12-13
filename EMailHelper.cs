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
        public static void SendMailServer(string title, string target, string text, bool isHtml)
        {
            SmtpClient mail = new SmtpClient(getXML("Host"),Convert.ToInt32(getXML("port")));
            string userName = getXML("userName");
            mail.UseDefaultCredentials = false;
            mail.Credentials = new NetworkCredential(userName + "@microsoft.com", getXML("FromPwd"));
            mail.EnableSsl = true;
            string subject = string.Format("{0}:{1}", title, CommonHelper.GetNowString());
            MailMessage msg = new MailMessage();
            //msg.From = new MailAddress(System.Environment.UserName + "@microsoft.com");
            msg.From = new MailAddress(userName + "@microsoft.com");
            msg.Subject = subject;
            List<string> tos = ParseToTarget(target);
            foreach (string to in tos)
            {
                msg.To.Add(to);
            }
            msg.IsBodyHtml = true;
            msg.Body = text;
            try
            {
                mail.Send(msg);
            }
            catch (Exception ex)
            {
                //LogHelper.DumpLog("SendMail Something is wrong", ex.ToString(),"log",DateTime.Now);
                //SendMailToDev(ex.ToString());
            }
        }

        public static string getXML(string nodeName)
        {
            string strReturn = "";
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load("./Config.xml");
            //根据路径获取节点
            XmlNode xmlNode = null;
            xmlNode = xdoc.SelectSingleNode("configurationN/" + nodeName);
            strReturn = xmlNode.InnerText;
            return strReturn;
        }


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
