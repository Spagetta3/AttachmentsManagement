using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Common;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Security;
using System.Net;


namespace AttachmentsManagementAgent
{
    public class AttachmentsManagementAgentFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            RoutingAgent myAgent = new AttachmentsManagementAgent();
            return myAgent;
        }
    }

    public class AttachmentsManagementAgent : RoutingAgent
    {
        private string domain;
        private int minSizeOfAttachment;
        private string login;
        private string password;
        private List<Microsoft.Exchange.Data.Transport.Email.Attachment> attachmentsToManage;
        private List<EnvelopeRecipient> recipientsOutsideOfCompany;

        public AttachmentsManagementAgent()
        {
            LoadConfig();
            attachmentsToManage = new List<Microsoft.Exchange.Data.Transport.Email.Attachment>();
            recipientsOutsideOfCompany = new List<EnvelopeRecipient>();
            base.OnResolvedMessage += new ResolvedMessageEventHandler(OnResolvedMessage);
        }

        private void OnResolvedMessage(ResolvedMessageEventSource source, QueuedMessageEventArgs e)
        {
            try
            {
                //Declare log string
                String NextLine = String.Empty;
                String delivery = e.MailItem.InboundDeliveryMethod.ToString();

                var senderDomain = e.MailItem.FromAddress.DomainPart;
                
                // Do not manage attachments for external senders outside of company 
                if (senderDomain != domain)
                {
                    return;
                }

                // Do not manage attachments with lower size than minSizeOfAttachment attribute in config
                foreach (Microsoft.Exchange.Data.Transport.Email.Attachment atAttach in e.MailItem.Message.Attachments)
                {
                    if (atAttach.AttachmentType == AttachmentType.Regular & atAttach.FileName != null)
                    {
                        Stream attachstream = atAttach.GetContentReadStream();
                        if (attachstream.Length >= (minSizeOfAttachment*1024*1024))
                        {
                            attachmentsToManage.Add(atAttach);
                        }
                    }
                }

                if (attachmentsToManage.Count > 0)
                {
                    var targetSite = new Uri("http://mysharepoint");
                    NetworkCredential cred = new NetworkCredential(login, password, domain);

                    using (ClientContext clientContext = new ClientContext(targetSite))
                    {
                        clientContext.Credentials = cred;
                        Web web = clientContext.Web;
                        clientContext.Load(web,
                        webSite => webSite.Title);

                        clientContext.ExecuteQuery();

                        NextLine = "Title is: " + web.Title;
                    }

                    NextLine  += "     ;Count of attachments: " + attachmentsToManage.Count;

                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                    {
                        file.WriteLine(NextLine);
                    }

                    return;

                    //String fileToUpload = @"C:\YourFile.txt";
                    //String sharePointSite = "http://yoursite.com/sites/Research/";
                    //String documentLibraryName = "Shared Documents";

                    //using (SPSite oSite = new SPSite(sharePointSite))
                    //{
                    //    using (SPWeb oWeb = oSite.OpenWeb())
                    //    {
                    //        if (!System.IO.File.Exists(fileToUpload))
                    //            throw new FileNotFoundException("File not found.", fileToUpload);

                    //        SPFolder myLibrary = oWeb.Folders[documentLibraryName];

                    //        // Prepare to upload
                    //        Boolean replaceExistingFiles = true;
                    //        String fileName = System.IO.Path.GetFileName(fileToUpload);
                    //        FileStream fileStream = File.OpenRead(fileToUpload);

                    //        // Upload document
                    //        SPFile spfile = myLibrary.Files.Add(fileName, fileStream, replaceExistingFiles);

                    //        // Commit 
                    //        myLibrary.Update();
                    //    }
                    //}
                }

                //Go through the recipients and check which ones are NOT internal-> Overwrite their routing
                foreach (EnvelopeRecipient recp in e.MailItem.Recipients)
                {
                    //Do internal check (same domain)       
                    if (recp.Address.DomainPart.Equals(senderDomain, StringComparison.OrdinalIgnoreCase)) //  if (!recp.Address.DomainPart.Equals(senderDomain, StringComparison.OrdinalIgnoreCase))
                    {
                        //This isn't an internal email. Trying to do sth.
                        try
                        {
                            e.MailItem.Message.Subject += " - FUCK YOU DID IT :P";
                        }
                        catch (Exception ex)
                        {
                            NextLine = "Failed to override setting:" + ex;

                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                            {
                                file.WriteLine(NextLine);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                String NextLine = "I'm inside the main catch statement" + ex;
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                {
                    file.WriteLine(NextLine);
                }
            }
        }

        private void LoadConfig()
        {
            try
            {
                string configFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml";

                XmlDocument doc = new XmlDocument();
                doc.Load(configFile);

                XmlNodeList nodes = doc.SelectNodes("configuration/key");
                foreach (XmlNode node in nodes)
                {
                    XmlAttributeCollection nodeAtt = node.Attributes;

                    foreach (XmlAttribute att in nodeAtt)
                    {
                        XmlDocument childNode = new XmlDocument();
                        childNode.LoadXml(node.OuterXml);

                        switch (att.Value)
                        {
                            case "domain":
                                domain = childNode.SelectSingleNode("key/value").InnerText;
                                break;
                            case "minSizeOfAttachment":
                                minSizeOfAttachment = int.Parse(childNode.SelectSingleNode("key/value").InnerText);
                                break;
                            case "login":
                                login = childNode.SelectSingleNode("key/value").InnerText;
                                break;
                            case "password":
                                password = childNode.SelectSingleNode("key/value").InnerText;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String NextLine = "Failed to load config file";
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"c:\temp\transportLogs.txt", true))
                {
                    file.WriteLine(NextLine);
                }
            }
        }
    }
}