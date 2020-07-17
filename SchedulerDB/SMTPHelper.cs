using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SqlServerHelper
{
    public class SMTPHelper
    {
        private string _from { get; set; }
        private string _password { get; set; }
        private string _host { get; set; }
        private int _port { get; set; }
        private bool _enableSSL { get; set; }
        private bool _isUseDefaultCredentials { get; set; }

        private SmtpClient _smtp { get; set; }

        public SMTPHelper(string from, string password, string host, int port, bool enableSSL, bool isUseDefaultCredentials)
        {
            _from = from;
            _password = password;
            _host = host;
            _port = port;
            _enableSSL = enableSSL;
            _isUseDefaultCredentials = isUseDefaultCredentials;

            _smtp = new SmtpClient
            {
                Host = _host, //"smtp.gmail.com",
                Port = _port, //587,
                EnableSsl = _enableSSL, //true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = _isUseDefaultCredentials, //false,
                Credentials = new NetworkCredential(_from, _password),
            };
        }

        public SMTPHelper(string from, string password, string host, int port)
        {
            _from = from;
            _password = password;
            _host = host;
            _port = port;
            _enableSSL = true;
            _isUseDefaultCredentials = false;

            _smtp = new SmtpClient
            {
                Host = _host, //"smtp.gmail.com",
                Port = _port, //587,
                EnableSsl = _enableSSL, //true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = _isUseDefaultCredentials, //false,
                Credentials = new NetworkCredential(_from, _password),
            };
        }

        public SMTPHelper(string from, string host, int port, bool enableSSL, bool isUseDefaultCredentials)
        {
            _from = from;

            _host = host;
            _port = port;
            _enableSSL = enableSSL;
            _isUseDefaultCredentials = isUseDefaultCredentials;

            _smtp = new SmtpClient
            {
                Host = _host, //"smtp.gmail.com",
                Port = _port, //587,
                EnableSsl = _enableSSL, //true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = _isUseDefaultCredentials, //false,
                Credentials = CredentialCache.DefaultNetworkCredentials,
            };
        }

        public void SendMail(string to, string cc, string bcc, string subject, string body, params string[] attachmentList)
        {
            var fromAddress = new MailAddress(_from);
            //var toAddress = new MailAddress("tyen@microsoft.com", "Leon Yen(MS)");

            try
            {
                using (var message = new MailMessage()
                {
                    Subject = subject,
                    Body = body.Replace("\r\n", "<br>"),
                    From = fromAddress,
                    IsBodyHtml = true
                    //mail.Attachments.Add(new Attachment("C:\\file.zip"));
                })
                {
                    if (!string.IsNullOrEmpty(to))
                    {
                        foreach (string addr in to.Split(';'))
                        {
                            message.To.Add(new MailAddress(addr));
                        }
                    }

                    if (!string.IsNullOrEmpty(cc))
                    {
                        foreach (string addr in cc.Split(';'))
                        {
                            message.CC.Add(new MailAddress(addr));
                        }
                    }

                    if (!string.IsNullOrEmpty(bcc))
                    {
                        foreach (string addr in bcc.Split(';'))
                        {
                            message.Bcc.Add(new MailAddress(addr));
                        }
                    }

                    //message.To.Add(new MailAddress("blackjackarchi@dualred.onmicrosoft.com", "Black Jack Archi(MS)"));
                    //message.To.Add(new MailAddress("blackjackdevlead@dualred.onmicrosoft.com", "Black Jack Dev Lead(MS)"));
                    //message.To.Add(new MailAddress("blackjackdb@dualred.onmicrosoft.com", "Black Jack DB(MS)"));
                    //message.To.Add(new MailAddress("blackjackdev8@microsoft.com", "Black Jack Dev8(MS)"));

                    //message.CC.Add(new MailAddress("tyen@microsoft.com", "Leon Yen(MS)"));
                    //message.CC.Add(new MailAddress("geoliang@microsoft.com", "George Liang(MS)"));
                    //message.CC.Add(new MailAddress("Sky.Hung@microsoft.com", "Sky Hung(MS)"));
                    //message.CC.Add(new MailAddress("Sol.Lee@microsoft.com", "Sol Lee(MS)"));
                    //message.CC.Add(new MailAddress("v-jahsi@microsoft.com", "Jason Hsieh(MS)"));
                    //message.CC.Add(new MailAddress("v-cjian@microsoft.com", "Chen Jian(MS)"));

                    if (attachmentList != null && attachmentList.Length > 0)
                    {
                        foreach (string attachment in attachmentList)
                        {
                            message.Attachments.Add(new Attachment(attachment));
                        }
                    }

                    _smtp.Send(message);
                }
            }
            catch (Exception ex)
            {
                SendMail("lovepudding0420@outlook.com", ex);
            }

        }

        public void SendMail(string to, Exception ex)
        {
            using (var message = new MailMessage() { Subject = $"Exception", Body = $"{ex.Message}", From = new MailAddress(_from) })
            {
                if (!string.IsNullOrEmpty(to))
                {
                    foreach (string addr in to.Split(';'))
                    {
                        message.To.Add(new MailAddress(addr));
                    }
                }

                _smtp.Send(message);
            }
        }

        //SMTPHelper helper1 = new SMTPHelper("meteors.sky@outlook.com",
        //                           "",
        //                           "smtp.office365.com",
        //                           Properties.Settings.Default.SMTPPort,
        //                           Properties.Settings.Default.SMTPEnableSsl,
        //                           Properties.Settings.Default.UseDefaultCredentials);

        //string subject1 = $"Hello world";
        //string body1 = $"Hi All, Hello Kitty，";

        //List<string> attachments1 = new List<string>();

        //helper1.SendMail("tyen@microsoft.com", null, null, subject1, body1, attachments1.ToArray());
    }
}
