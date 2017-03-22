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
        private List<Attachment> attachmentsToManage;
        private List<EnvelopeRecipient> recipientsOutsideOfCompany;

        public AttachmentsManagementAgent()
        {
            LoadConfig();
            attachmentsToManage = new List<Attachment>();
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
                foreach (Attachment atAttach in e.MailItem.Message.Attachments)
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