
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
      static ShoutMail()
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
            //new ShoutMail();
            // POP3_Check_Msg();
             search_mail();
            // ReadImap();
            //Imap_Check_Msg();

        }
     
       
        
        static IMailFolder get_mail()
        {
           
            IMailFolder inbox_data=null;
            try
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.Connect(Constants.mail_server, Constants.port, Constants.use_SSL);
                client.Authenticate(uname,password);
              var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);
               inbox_data = inbox;
               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return inbox_data;
        }
        static void search_mail()
        {
            try
            {
                var inbox = get_mail();
                foreach (var uid in inbox.Search(SearchQuery.NotSeen.And(SearchQuery.DeliveredOn(DateTime.Today))))
                {
                    var message = inbox.GetMessage(uid);
                    Display_Mail(message, uid);
                    inbox.AddFlags(uid, MessageFlags.Seen, true);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                client.Disconnect(true);
            }
        }
        static void Display_Mail(MimeMessage message,UniqueId uid) {
            try
            {
                Console.WriteLine("[match] {0}: {1}", uid, message.Subject);
                //MessageBox.Show(message.Subject.ToString());
                Speak_Mail(message);
            }
            catch (Exception e)
            {
               MessageBox.Show(e.Message);
            }
        }
  static void Speak_Mail(MimeMessage message)
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

