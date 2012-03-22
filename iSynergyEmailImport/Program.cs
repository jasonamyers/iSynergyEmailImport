using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Text.RegularExpressions;
using Email.Net.Common.Configurations;
using Email.Net.Common.MessageParts;
using Email.Net.Imap.Responses;
using Email.Net.Imap.Collections;
using Email.Net.Imap;
using Email.Net.Imap.Sequences;


namespace iSynergyEmailImport
{
    class Program
    {
        static void Main(string[] args)
        {
            ImapClient client = new Email.Net.Imap.ImapClient();
            client.Host = "host";
            client.Username = "User";
            client.Password = "password";
            client.Port = (ushort)993;
            client.AttachmentDirectory = "PathForAttachments";
            client.SSLInteractionType = EInteractionType.SSLPort;

            Mailbox inbox;
            ImapMessage msg;
            CompletionResponse response = client.Login();
            if (response.CompletionResult == ECompletionResponseType.OK)
            {
                Console.WriteLine("We're In");
                Mailbox folders = client.GetMailboxTree();
                inbox = Mailbox.Find(folders, "INBOX");
                Console.WriteLine("{0} Messages", client.GetMessageCount(inbox));
                MessageCollection tmp = client.GetAllMessageHeaders(inbox);
                foreach(ImapMessage x in tmp)
                {
                    msg = client.GetMessageText(x.UID, inbox);
                    if (msg.Attachments.Count > 0)
                    {
                        for (int i = 0; i <msg.Attachments.Count; i++)
                        {
                            Attachment attach = msg.Attachments[i];
                            if (attach != null)
                            {
                                Attachment tmpAttach = new Attachment();
                                
                                tmpAttach = client.GetAttachment(inbox, msg, msg.Attachments[i]);
                                string tmpName = tmpAttach.FullFilename.ToString();
                                string realName = client.AttachmentDirectory + msg.Attachments[i].TransferFilename.ToString();
                                File.Move(tmpName, realName);
                                Console.WriteLine("Attachment: {0}", realName);
                            }
                        }
                        ISequence squence = new SequenceNumber(x.UID);
                        CompletionResponse delResponse = client.MarkAsDeleted(squence, inbox);
                        delResponse = client.DeleteMarkedMessages(inbox);
                    }
                }

            }
        }
    }
}
