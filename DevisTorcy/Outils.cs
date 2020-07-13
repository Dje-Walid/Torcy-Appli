using System;
using IniParser;
using IniParser.Model;
using System.Windows.Forms;
using System.Data.OleDb;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Net.Mail;
using Message = Google.Apis.Gmail.v1.Data.Message;

namespace DevisTorcy
{
    class Outils
    {
        private OleDbConnection connection;
        private FileIniDataParser config;
        string[] Scope = { GmailService.Scope.GmailSend };
        string ApplicationName;


        //Constructeur classe
        public Outils()
        {
            this.connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Torcy.accdb; Persist Security Info=False;");
            this.config = new FileIniDataParser();
            this.ApplicationName = "DevisTorcy";
        }

        #region OLEDB ACCESS
        public void setConnection(string NomFichier)
        {
            this.connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\" + NomFichier + ".accdb; Persist Security Info=False;");
        }

        public OleDbConnection getConnection()
        {
            return this.connection;
        }

        public void testConnection()
        {
            try
            {
                this.connection.Open();
                MessageBox.Show("Connexion fonctionne");
                this.connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connexion ne fonctionne pas" + ex);
            }
        }
        #endregion

        #region INI PARSER
        public IniData getConfig()
        {
            IniData data = this.config.ReadFile("Config.ini");
            return data;
        }

        public int getNumDevis()
        {
            return Convert.ToInt32(this.getConfig()["Info"]["NumDevis"]);
        }

        public void setNumDevis(string num)
        {
            IniData data = this.getConfig();
            data["Info"]["NumDevis"] = num;
            this.config.WriteFile("Config.ini", data);
        }

        public int getNumFacture()
        {
            return Convert.ToInt32(this.getConfig()["Info"]["NumFacture"]);
        }

        public void setNumFacture(string num)
        {
            IniData data = this.getConfig();
            data["Info"]["NumFacture"] = num;
            this.config.WriteFile("Config.ini", data);
        }
        #endregion

        #region GMAIL
        private string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }
        public void sendMail(string mailClient, string nomFichier)
        {
            //Chemin d'accès
            DirectoryInfo dirBeforeAppli = Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Directory.GetCurrentDirectory())))))));

            //Create Message
            MailMessage mail = new MailMessage();
            mail.Subject = "Devis Île de Loisirs de Vaires-Torcy";
            mail.Body = "Bonjour, \r\n\r\nVeuillez trouver ci joint votre devis pour la baignade. \r\n\r\nBonne journee, Cordialement. \nDJEMMAA Walid\r\n\r\n\r\n--\r\nAccueil Plage\r\naccueilplage@vaires-torcy.iledeloisirs.fr\r\nPole baignade\r\nRoute de Lagny\r\n77200 Torcy\r\nTel: 0160200204 (touche 2)\r\nFax: 0164809149";
            mail.From = new MailAddress("accueilplage@vaires-torcy.iledeloisirs.fr");
            mail.IsBodyHtml = false;
            string joint = dirBeforeAppli + @"\Devis" + DateTime.Today.Year.ToString() + @"\5860" + (DateTime.Today.Year - 2000) + "-" + nomFichier + ".xlsx";
            mail.Attachments.Add(new Attachment(joint));
            mail.To.Add(new MailAddress(mailClient));
            MimeKit.MimeMessage mimeMessage = MimeKit.MimeMessage.CreateFromMailMessage(mail);

            Google.Apis.Gmail.v1.Data.Message message = new Message();
            message.Raw = Base64UrlEncode(mimeMessage.ToString());

            //Gmail API credentials
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart2.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scope,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            //Send Email
            var result = service.Users.Messages.Send(message,"me").Execute();
        }
        #endregion
    }
}
