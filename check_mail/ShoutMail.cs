
using System;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.Windows;
using System.Speech.Synthesis;
using ActiveUp.Net.Mail;
using MimePart = MimeKit.MimePart;

namespace check_mail
{
    class ShoutMail
    {
        private static MailKit.Net.Imap.ImapClient client;
        private static string uname;
        private static string password;
        private static string mail_server;
        public static int port;
        public static bool use_SSL;
        ShoutMail()
        {
            client = new MailKit.Net.Imap.ImapClient();
            uname = Constants.uname;
            password = Constants.password1;
            mail_server = Constants.mail_server;
            port = Constants.port;
            use_SSL = Constants.use_SSL;

        }


        static void Main(string[] args)
        {
            // POP3_Check_Msg();
            // search_mail();
            // ReadImap();
            Imap_Check_Msg();

        }
        public static void ReadImap()
        {
            var mailRepository = new MailRepository(
                                    "imap.gmail.com",
                                    993,
                                    true,
                                    "eartherk5yb@gmail.com",
                                    "ztlyyzguwntwwrwq"
                                );

            var emailList = mailRepository.GetAllMails("inbox");

            foreach (Message email in emailList)
            {
                Console.WriteLine("<p>{0}: {1}</p><p>{2}</p>", email.From, email.Subject, email.BodyHtml.Text);
                Console.Read();
                if (email.Attachments.Count > 0)
                {
                    foreach (MimePart attachment in email.Attachments)
                    {
                        Console.WriteLine("<p>Attachment: {0} {1}</p>", attachment.ContentBase, attachment.ContentType.MimeType);
                    }
                }
            }
        }
        static void Imap_Check_Msg() {

            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                // For demo-purposes, accept all SSL certificates
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                client.Connect("imap.gmail.com", 993, true);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate("bhargavak37@gmail.com", "raxlgsnjjbkjcqkv");
               // The Inbox folder is always available on all IMAP servers...
               var inbox = client.Inbox;
               inbox.Open(FolderAccess.ReadOnly);

              Console.WriteLine("Total messages: {0}", inbox.Count);
                Console.WriteLine("Recent messages: {0}", inbox.Recent);

                for (int i = 0; i < inbox.Count; i++)
                {
                    var message = inbox.GetMessage(i);
                    Console.WriteLine("Subject: {0}", message.Subject);
                    //MessageBox.Show(message.Subject);
                }

                client.Disconnect(true);
            }
        }
        //static void POP3_Check_Msg() {
        //    using (var client = new Pop3Client())
        //    {
        //        // For demo-purposes, accept all SSL certificates (in case the server supports STARTTLS)
        //        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        //        client.Connect("pop.gmail.com", 995, true);

        //        // Note: since we don't have an OAuth2 token, disable
        //        // the XOAUTH2 authentication mechanism.
        //        client.AuthenticationMechanisms.Remove("XOAUTH2");

        //        client.Authenticate(uname, password);

        //        for (int i = 0; i < client.Count; i++)
        //        {
        //            var message = client.GetMessage(i);
        //            Console.WriteLine("Subject: {0}", message.Subject);
        //            MessageBox.Show(message.Subject);
        //        }

        //        client.Disconnect(true);
        //    }
        //}
        static void search_mail()
        {
            var client = new MailKit.Net.Imap.ImapClient();
            try
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(Constants.mail_server, Constants.port, Constants.use_SSL);
                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                //client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(Constants.uname, Constants.password1);
                // The Inbox folder is always available on all IMAP servers...
                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);
                //var query = SearchQuery.DeliveredAfter(DateTime.Parse("2017-06-28"))
                //.And(SearchQuery.SubjectContains("*"))
                //.And(SearchQuery.Answered);
                Console.Write(DateTime.Today);
                foreach (var uid in inbox.Search(SearchQuery.NotSeen.And(SearchQuery.DeliveredOn(DateTime.Today))))
                {
                    var message = inbox.GetMessage(uid);
                    Console.WriteLine("[match] {0}: {1}", uid, message.Subject);
                    MessageBox.Show(message.Subject.ToString());
                    Shout_Mail(message);
                    inbox.AddFlags(uid, MessageFlags.Seen, true);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            finally {
                client.Disconnect(true);
            }
            /*  foreach (var summary in inbox.Fetch(0, 7, MessageSummaryItems.Full | MessageSummaryItems.UniqueId))
              {
                  Console.WriteLine("[summary] {0:D2}: {1}", summary.Index, summary.Envelope.Subject);
                  // MessageBox.Show(summary.Envelope.Subject.ToString());
              }*/

        }
  static void Shout_Mail(MimeMessage message)
        {
            using (SpeechSynthesizer synth =
   new SpeechSynthesizer())
            {
                synth.Volume = 100; synth.Rate = -2;
                synth.SelectVoiceByHints(VoiceGender.Male);
                synth.Speak(message.Subject);
            }
        }

    }
}

