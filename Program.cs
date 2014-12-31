using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;

namespace SendMailViaExchange
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback =
                (sender, certificate, chain, sslPolicyErrors) =>
                {
                    if (sslPolicyErrors == SslPolicyErrors.None) return true;
                    if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == 0) return false;
                    if (chain == null || chain.ChainStatus == null) return true;
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                            (status.Status == X509ChainStatusFlags.UntrustedRoot))
                            continue;
                        if (status.Status != X509ChainStatusFlags.NoError) return false;
                    }
                    return true;
                };
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            //service.Url = new Uri("https://aps.mail.microsoft.com/ews/Services.wsdl");
            service.Credentials = new WebCredentials("username@servername.com", "password");
            service.UseDefaultCredentials = true;
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
            service.AutodiscoverUrl("username@servername.com", url => (new Uri(url).Scheme == "https"));
            EmailMessage mail = new EmailMessage(service);
            mail.ToRecipients.Add("username@servername.com");
            mail.Subject = "HelloWorld";
            mail.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            mail.Send();
        }
    }
}

