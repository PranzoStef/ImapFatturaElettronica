using MailKit;
using MailKit.Search;
using MimeKit;
using System;
using System.IO;
//OpenPop.NET
namespace ImapFatturaElettronica
{
    class Program
    {
        static void Main(string[] args)
        {
            //variabili comuni
            string imapServer = "imaps.pec.aruba.it";
            string imapPort = "993";
            string emailAddress = "pec@pec.it";
            string emailPass = "xxxxxxxxx"; 
            string ID_UltimaFattura = "0";
            string DataUltimoControlloFatture = "01/01/2020";
            string folderXML = "C:\\XmlFE\\";

            // se la cartella non esiste, la creo
            if(!System.IO.Directory.Exists(folderXML))
            {
                System.IO.Directory.CreateDirectory(folderXML);
            }

            MailKit.Net.Imap.ImapClient client = new MailKit.Net.Imap.ImapClient();
            client.CheckCertificateRevocation = false;
            client.Connect(imapServer, int.Parse(imapPort), true);
            client.Authenticate(emailAddress, emailPass);
            client.Inbox.Open(FolderAccess.ReadOnly);

            int ultimouid = int.Parse(ID_UltimaFattura);
            var query = SearchQuery.DeliveredAfter(DateTime.Parse(DataUltimoControlloFatture))
                    .And(SearchQuery.FromContains("sdi"));
            foreach (var item in client.Inbox.Search(query))
            {
                if (item.Id > ultimouid)
                {
                    MimeKit.MimeMessage messaggio = client.Inbox.GetMessage(item);

                    foreach (var bodyPart in messaggio.BodyParts)
                    {
                        if (bodyPart.ContentDisposition.FileName == "postacert.eml")
                        {
                            MimeKit.MessagePart _postacert = (MessagePart)bodyPart;
                            if (_postacert.Message.Subject.Contains("Invio File"))
                            {
                                foreach (var attachment in _postacert.Message.Attachments)
                                {
                                    if (attachment is MessagePart)
                                    {
                                        var fileName = attachment.ContentDisposition?.FileName;
                                        var rfc822 = (MessagePart)attachment;

                                        if (string.IsNullOrEmpty(fileName))
                                            fileName = "attached-message.eml";

                                        using (var stream = File.Create(fileName))
                                            rfc822.Message.WriteTo(stream);
                                    }
                                    else
                                    {
                                        var part = (MimePart)attachment;
                                        var fileName = part.FileName;
                                        // verifico se il file esiste o meno.
                                        if(!System.IO.File.Exists(folderXML + fileName))
                                        {
                                            FileStream file = new FileStream(folderXML + fileName, FileMode.Create);
                                            part.Content.DecodeTo(file);
                                            file.Close();
                                        }                                       
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
    }
}
