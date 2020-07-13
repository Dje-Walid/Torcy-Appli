using DevisTorcy.Forms;
using System;
using System.IO;
using System.Data.OleDb;
using System.Windows.Forms;

namespace DevisTorcy
{
    public partial class Devis : Form
    {
        int nbDate = 1;

        public Devis()
        {
            InitializeComponent();
        }

        private void Devis_Load(object sender, EventArgs e)
        {
            txbxNumDevis.Text = Convert.ToString(Program.outils.getNumDevis());

            //Remplissage cbxVille
            Program.outils.getConnection().Open();
            string requete = "Select [Ville] from DevisAdresse;";
            OleDbCommand cmd = new OleDbCommand(requete, Program.outils.getConnection());
            OleDbDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                cbxVille.Items.Add(dr[0].ToString());
            }
            Program.outils.getConnection().Close();
        }

        #region Menu
        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.Show();
        }

        private void factureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Facture facture = new Facture();
            facture.Show();
        }

        private void dPAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            DPAE dpae = new DPAE();
            dpae.Show();
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            switch (nbDate)
            {
                case 1:
                    grbDate2.Visible = true;
                    nbDate += 1;
                    break;
                case 2:
                    grbDate3.Visible = true;
                    nbDate += 1;
                    break;
                case 3:
                    grbDate4.Visible = true;
                    nbDate += 1;
                    break;
                case 4:
                    grbDate5.Visible = true;
                    nbDate += 1;
                    break;
                case 5:
                    grbDate6.Visible = true;
                    nbDate += 1;
                    break;
                case 6:
                    MessageBox.Show("Pour le moment vous êtes limiter à 6 dates par devis, une des solutions serait de faire le devis manuellement.", "Limite de Date :", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
            }
        }

        private void btSuppDate_Click(object sender, EventArgs e)
        {
            switch (nbDate)
            {
                case 1:
                    MessageBox.Show("Afin d'établir le devis vous devez au moins renseigner une date.", "Date minimum :", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case 2:
                    grbDate2.Visible = false;
                    nbDate -= 1;
                    break;
                case 3:
                    grbDate3.Visible = false;
                    nbDate -= 1;
                    break;
                case 4:
                    grbDate4.Visible = false;
                    nbDate -= 1;
                    break;
                case 5:
                    grbDate5.Visible = false;
                    nbDate -= 1;
                    break;
                case 6:
                    grbDate6.Visible = false;
                    nbDate -= 1;
                    break;
            }

        }

        private void btSend_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Directory.GetCurrentDirectory() + @"\ExempleDevis.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            DirectoryInfo dirBeforeAppli = Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Directory.GetCurrentDirectory())))))));
            Directory.CreateDirectory(dirBeforeAppli + @"\Devis" + DateTime.Today.Year.ToString());

            //Compter le nombre de lignes
            Microsoft.Office.Interop.Excel.Range userRange = x.UsedRange;
            int countRecords = userRange.Rows.Count;
            int add = countRecords + 1;
            x.Cells[add, 1] = "Ligne Total " + countRecords;

            //Mettre en gras une ligne
            //x.Range["A1"].EntireRow.Font.Bold = true;
            //Mettre en gras une cell
            //x.Range["A1"].Font.Bold = true;

            //Num devis +1 et Excel
            x.Range["F3"].Value = Program.outils.getNumDevis();
            Program.outils.setNumDevis(Convert.ToString(Program.outils.getNumDevis() + 1));

            //Date du jour
            x.Range["E4"].Value = DateTime.Today.ToShortDateString();

            //Adresse
            Program.outils.getConnection().Open();
            string requete = "Select [LigneAdr1] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete1 = "Select [LigneAdr2] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete2 = "Select [LigneAdr3] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete3 = "Select [LigneAdr4] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete4 = "Select [CodePostal] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete5 = "Select [Ville] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";

            OleDbCommand cmd = new OleDbCommand(requete, Program.outils.getConnection());
            OleDbDataReader dr = cmd.ExecuteReader();
            OleDbCommand cmd1 = new OleDbCommand(requete1, Program.outils.getConnection());
            OleDbDataReader dr1 = cmd1.ExecuteReader();
            OleDbCommand cmd2 = new OleDbCommand(requete2, Program.outils.getConnection());
            OleDbDataReader dr2 = cmd2.ExecuteReader();
            OleDbCommand cmd3 = new OleDbCommand(requete3, Program.outils.getConnection());
            OleDbDataReader dr3 = cmd3.ExecuteReader();
            OleDbCommand cmd4 = new OleDbCommand(requete4, Program.outils.getConnection());
            OleDbDataReader dr4 = cmd4.ExecuteReader();
            OleDbCommand cmd5 = new OleDbCommand(requete5, Program.outils.getConnection());
            OleDbDataReader dr5 = cmd5.ExecuteReader();


            while (dr.Read() && dr1.Read() && dr2.Read() && dr3.Read() && dr4.Read() && dr5.Read())
            {
                if (dr2[0].ToString() == "" && dr3[0].ToString() == "")
                {
                    x.Range["D8"].Value = dr[0].ToString();
                    x.Range["D9"].Value = dr1[0].ToString();
                    x.Range["D10"].Value = dr4[0].ToString() + " " + dr5[0].ToString();
                }
                else if (dr3[0].ToString() == "")
                {
                    x.Range["D8"].Value = dr[0].ToString();
                    x.Range["D9"].Value = dr1[0].ToString();
                    x.Range["D10"].Value = dr2[0].ToString();
                    x.Range["D11"].Value = dr4[0].ToString() + " " + dr5[0].ToString();
                }
                else if (dr3[0].ToString() != "")
                {
                    x.Range["D7"].Font.Bold = true;
                    x.Range["D8"].Font.Bold = false;
                    x.Range["D7"].Value = dr[0].ToString();
                    x.Range["D8"].Value = dr1[0].ToString();
                    x.Range["D9"].Value = dr2[0].ToString();
                    x.Range["D10"].Value = dr3[0].ToString();
                    x.Range["D11"].Value = dr4[0].ToString() + " " + dr5[0].ToString();
                }
            }
            Program.outils.getConnection().Close();

            switch (nbDate)
            {
                case 1:
                    #region Suppression lignes en trop
                    //Suppression des lignes en trop
                    x.Range["A19"].Value = "";
                    x.Range["D19"].Value = "";
                    x.Range["D20"].Value = "";
                    x.Range["D20"].Value = "";
                    x.Range["B19"].Value = "";
                    x.Range["B20"].Value = "";
                    x.Range["E19"].Value = "";
                    x.Range["E20"].Value = "";

                    x.Range["A22"].Value = "";
                    x.Range["D22"].Value = "";
                    x.Range["D23"].Value = "";
                    x.Range["B22"].Value = "";
                    x.Range["B23"].Value = "";
                    x.Range["E22"].Value = "";
                    x.Range["E23"].Value = "";

                    x.Range["A25"].Value = "";
                    x.Range["D25"].Value = "";
                    x.Range["D26"].Value = "";
                    x.Range["B25"].Value = "";
                    x.Range["B26"].Value = "";
                    x.Range["E25"].Value = "";
                    x.Range["E26"].Value = "";

                    x.Range["A28"].Value = "";
                    x.Range["D28"].Value = "";
                    x.Range["D29"].Value = "";
                    x.Range["B28"].Value = "";
                    x.Range["B29"].Value = "";
                    x.Range["E28"].Value = "";
                    x.Range["E29"].Value = "";

                    x.Range["A31"].Value = "";
                    x.Range["D31"].Value = "";
                    x.Range["D32"].Value = "";
                    x.Range["B31"].Value = "";
                    x.Range["B32"].Value = "";
                    x.Range["E31"].Value = "";
                    x.Range["E32"].Value = "";
                    #endregion
                    break;
                case 2:
                    #region Suppression lignes en trop
                    //Suppression des lignes en trop
                    x.Range["A22"].Value = "";
                    x.Range["D22"].Value = "";
                    x.Range["D23"].Value = "";
                    x.Range["B22"].Value = "";
                    x.Range["B23"].Value = "";
                    x.Range["E22"].Value = "";
                    x.Range["E23"].Value = "";

                    x.Range["A25"].Value = "";
                    x.Range["D25"].Value = "";
                    x.Range["D26"].Value = "";
                    x.Range["B25"].Value = "";
                    x.Range["B26"].Value = "";
                    x.Range["E25"].Value = "";
                    x.Range["E26"].Value = "";

                    x.Range["A28"].Value = "";
                    x.Range["D28"].Value = "";
                    x.Range["D29"].Value = "";
                    x.Range["B28"].Value = "";
                    x.Range["B29"].Value = "";
                    x.Range["E28"].Value = "";
                    x.Range["E29"].Value = "";

                    x.Range["A31"].Value = "";
                    x.Range["D31"].Value = "";
                    x.Range["D32"].Value = "";
                    x.Range["B31"].Value = "";
                    x.Range["B32"].Value = "";
                    x.Range["E31"].Value = "";
                    x.Range["E32"].Value = "";
                    #endregion
                    break;
                case 3:
                    #region Suppression lignes en trop
                    //Suppression des lignes en trop
                    x.Range["A25"].Value = "";
                    x.Range["D25"].Value = "";
                    x.Range["D26"].Value = "";
                    x.Range["B25"].Value = "";
                    x.Range["B26"].Value = "";
                    x.Range["E25"].Value = "";
                    x.Range["E26"].Value = "";

                    x.Range["A28"].Value = "";
                    x.Range["D28"].Value = "";
                    x.Range["D29"].Value = "";
                    x.Range["B28"].Value = "";
                    x.Range["B29"].Value = "";
                    x.Range["E28"].Value = "";
                    x.Range["E29"].Value = "";

                    x.Range["A31"].Value = "";
                    x.Range["D31"].Value = "";
                    x.Range["D32"].Value = "";
                    x.Range["B31"].Value = "";
                    x.Range["B32"].Value = "";
                    x.Range["E31"].Value = "";
                    x.Range["E32"].Value = "";
                    #endregion
                    break;
                case 4:
                    #region Suppression lignes en trop
                    //Suppression des lignes en trop
                    x.Range["A28"].Value = "";
                    x.Range["D28"].Value = "";
                    x.Range["D29"].Value = "";
                    x.Range["B28"].Value = "";
                    x.Range["B29"].Value = "";
                    x.Range["E28"].Value = "";
                    x.Range["E29"].Value = "";

                    x.Range["A31"].Value = "";
                    x.Range["D31"].Value = "";
                    x.Range["D32"].Value = "";
                    x.Range["B31"].Value = "";
                    x.Range["B32"].Value = "";
                    x.Range["E31"].Value = "";
                    x.Range["E32"].Value = "";
                    #endregion
                    break;
                case 5:
                    #region Suppression lignes en trop
                    //Suppression des lignes en trop
                    x.Range["A31"].Value = "";
                    x.Range["D31"].Value = "";
                    x.Range["D32"].Value = "";
                    x.Range["B31"].Value = "";
                    x.Range["B32"].Value = "";
                    x.Range["E31"].Value = "";
                    x.Range["E32"].Value = "";
                    #endregion
                    break;
            }

            while (nbDate > 0)
            {
                switch (nbDate)
                {
                    case 1:
                        x.Range["A16"].Value = dtpDate1.Value.ToShortDateString();
                        x.Range["D16"].Value = txbxPlage1.Text;
                        x.Range["D17"].Value = txbxAccomp1.Text;

                        nbDate -= 1;
                        break;
                    case 2:
                        x.Range["A19"].Value = dtpDate2.Value.ToShortDateString();
                        x.Range["D19"].Value = txbxPlage2.Text;
                        x.Range["D20"].Value = txbxAccomp2.Text;

                        nbDate -= 1;
                        break;
                    case 3:
                        x.Range["A22"].Value = dtpDate3.Value.ToShortDateString();
                        x.Range["D22"].Value = txbxPlage3.Text;
                        x.Range["D23"].Value = txbxAccomp3.Text;

                        nbDate -= 1;
                        break;
                    case 4:
                        x.Range["A25"].Value = dtpDate4.Value.ToShortDateString();
                        x.Range["D25"].Value = txbxPlage4.Text;
                        x.Range["D26"].Value = txbxAccomp4.Text;

                        nbDate -= 1;
                        break;
                    case 5:
                        x.Range["A28"].Value = dtpDate5.Value.ToShortDateString();
                        x.Range["D28"].Value = txbxPlage5.Text;
                        x.Range["D29"].Value = txbxAccomp5.Text;

                        nbDate -= 1;
                        break;
                    case 6:
                        x.Range["A31"].Value = dtpDate6.Value.ToShortDateString();
                        x.Range["D31"].Value = txbxPlage6.Text;
                        x.Range["D32"].Value = txbxAccomp6.Text;
                        nbDate -= 1;
                        break;
                }
            }

            string nomFichier = dirBeforeAppli + @"\Devis" + DateTime.Today.Year.ToString() + @"\" + "VEUILLEZ INDIQUER LA VILLE" ;

            Program.outils.getConnection().Open();
            requete5 = "Select [Ville] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            cmd5.CommandText = requete5;
            dr5 = cmd5.ExecuteReader();

            while (dr5.Read())
            {
                nomFichier = dirBeforeAppli + @"\Devis" + DateTime.Today.Year.ToString() + @"\5860" + (DateTime.Today.Year - 2000) + "-" + Convert.ToString(Program.outils.getNumDevis() - 1) + " DEVIS " + dr5[0].ToString();
            }
            Program.outils.getConnection().Close();

            sheet.Close(true, nomFichier, Type.Missing);
            excel.Quit();

            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    p.Kill();
                }
            }

            //Envoi Mail
            if (txbxMail.Text != "")
            {
                Program.outils.sendMail(txbxMail.Text, Convert.ToString(Program.outils.getNumDevis() - 1) + " DEVIS " + cbxVille.Text);
            }

            this.Hide();
            Devis devis = new Devis();
            devis.Show();

        }

        private void btAddVille_Click(object sender, EventArgs e)
        {
            AjoutAdresseDevis addAdresse = new AjoutAdresseDevis();
            addAdresse.Show();
        }

        private void Devis_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void cbxVille_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Remplissage txbxAdresse
            Program.outils.getConnection().Open();
            string requete = "Select [LigneAdr1] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete1 = "Select [LigneAdr2] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete2 = "Select [LigneAdr3] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete3 = "Select [LigneAdr4] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete4 = "Select [CodePostal] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete5 = "Select [Ville] from DevisAdresse where [Ville] = \"" + cbxVille.Text + "\";";

            OleDbCommand cmd = new OleDbCommand(requete, Program.outils.getConnection());
            OleDbDataReader dr = cmd.ExecuteReader();
            OleDbCommand cmd1 = new OleDbCommand(requete1, Program.outils.getConnection());
            OleDbDataReader dr1 = cmd1.ExecuteReader();
            OleDbCommand cmd2 = new OleDbCommand(requete2, Program.outils.getConnection());
            OleDbDataReader dr2 = cmd2.ExecuteReader();
            OleDbCommand cmd3 = new OleDbCommand(requete3, Program.outils.getConnection());
            OleDbDataReader dr3 = cmd3.ExecuteReader();
            OleDbCommand cmd4 = new OleDbCommand(requete4, Program.outils.getConnection());
            OleDbDataReader dr4 = cmd4.ExecuteReader();
            OleDbCommand cmd5 = new OleDbCommand(requete5, Program.outils.getConnection());
            OleDbDataReader dr5 = cmd5.ExecuteReader();


            while (dr.Read() && dr1.Read() && dr2.Read() && dr3.Read() && dr4.Read() && dr5.Read())
            {
                if (dr2[0].ToString() == "" && dr3[0].ToString() == "")
                {
                    string s = dr[0].ToString();
                    string s1 = dr1[0].ToString();
                    string s2 = dr4[0].ToString();
                    string s3 = dr5[0].ToString();
                    txbxAdresse.Text = s + "\r\n" + s1 + "\r\n" + s2 + " " + s3;
                }
                else if (dr3[0].ToString() == "")
                {
                    string s = dr[0].ToString();
                    string s1 = dr1[0].ToString();
                    string s2 = dr2[0].ToString();
                    string s3 = dr4[0].ToString();
                    string s4 = dr5[0].ToString();
                    txbxAdresse.Text = s + "\r\n" + s1 + "\r\n" + s2 + "\r\n" + s3 + " " + s4;
                }
                else if (dr3[0].ToString() != "")
                {
                    string s = dr[0].ToString();
                    string s1 = dr1[0].ToString();
                    string s2 = dr2[0].ToString();
                    string s3 = dr3[0].ToString();
                    string s4 = dr4[0].ToString();
                    string s5 = dr5[0].ToString();
                    txbxAdresse.Text = s + "\r\n" + s1 + "\r\n" + s2 + "\r\n" + s3 + "\r\n" + s4 + " " + s5;
                }
            }
            Program.outils.getConnection().Close();
        }
    }
}
