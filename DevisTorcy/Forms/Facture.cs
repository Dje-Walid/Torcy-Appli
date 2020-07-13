using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using DevisTorcy.Forms;
using Microsoft.Office.Interop.Excel;

namespace DevisTorcy
{
    public partial class Facture : Form
    {
        int nbDate = 1;

        public Facture()
        {
            InitializeComponent();
        }

        #region Menu
        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Devis devis = new Devis();
            devis.Show();
        }

        private void factureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu menu = new Menu();
            menu.Show();
        }

        private void dPAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            DPAE dpae = new DPAE();
            dpae.Show();
        }
        #endregion

        private void Facture_Load(object sender, EventArgs e)
        {
            txbxNumFacture.Text = Convert.ToString(Program.outils.getNumFacture());

            //Remplissage cbxVille
            Program.outils.getConnection().Open();
            string requete = "Select [Ville] from FactureAdresse;";
            OleDbCommand cmd = new OleDbCommand(requete, Program.outils.getConnection());
            OleDbDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                cbxVille.Items.Add(dr[0].ToString());
            }
            Program.outils.getConnection().Close();
        }

        private void btSend_Click(object sender, EventArgs e)
        {
            if (txbxBDC.Text == "")
            {
                MessageBox.Show("Merci de remplir le numéro de bon de commande.", "Veuillez indiquer le numéro de BDC :", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Directory.GetCurrentDirectory() + @"\ExempleFacture.xls");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

                DirectoryInfo dirBeforeAppli = Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Convert.ToString(Directory.GetParent(Directory.GetCurrentDirectory())))))));
                Directory.CreateDirectory(dirBeforeAppli + @"\Facture" + DateTime.Today.Year.ToString());

                //Compter le nombre de lignes
                Microsoft.Office.Interop.Excel.Range userRange = x.UsedRange;
                int countRecords = userRange.Rows.Count;
                int add = countRecords + 1;
                x.Cells[add, 1] = "Ligne Total " + countRecords;

                //Mettre en gras une ligne
                //x.Range["A1"].EntireRow.Font.Bold = true;
                //Mettre en gras une cell
                //x.Range["A1"].Font.Bold = true;

                //Num Facture +1 et Excel
                if (Program.outils.getNumFacture() < 100)
                {
                    x.Range["F3"].Value = "0" + Program.outils.getNumFacture();
                }
                else
                {
                    x.Range["F3"].Value = Program.outils.getNumFacture();
                }
                Program.outils.setNumFacture(Convert.ToString(Program.outils.getNumFacture() + 1));

                //Numéro BDC
                x.Range["A6"].Value = "Selon votre bon de commande n° " + txbxBDC.Text;

                //Date du jour
                x.Range["E4"].Value = DateTime.Today.ToShortDateString();

                //Adresse
                Program.outils.getConnection().Open();
                string requete = "Select [LigneAdr1] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                string requete1 = "Select [LigneAdr2] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                string requete2 = "Select [LigneAdr3] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                string requete3 = "Select [LigneAdr4] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                string requete4 = "Select [CodePostal] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                string requete5 = "Select [Ville] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";

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
                        x.Range["A18"].Value = "";
                        x.Range["D18"].Value = "";
                        x.Range["D19"].Value = "";
                        x.Range["D19"].Value = "";
                        x.Range["B18"].Value = "";
                        x.Range["B19"].Value = "";
                        x.Range["E18"].Value = "";
                        x.Range["E19"].Value = "";

                        x.Range["A21"].Value = "";
                        x.Range["D21"].Value = "";
                        x.Range["D22"].Value = "";
                        x.Range["B21"].Value = "";
                        x.Range["B22"].Value = "";
                        x.Range["E21"].Value = "";
                        x.Range["E22"].Value = "";

                        x.Range["A24"].Value = "";
                        x.Range["D24"].Value = "";
                        x.Range["D25"].Value = "";
                        x.Range["B24"].Value = "";
                        x.Range["B25"].Value = "";
                        x.Range["E24"].Value = "";
                        x.Range["E25"].Value = "";

                        x.Range["A27"].Value = "";
                        x.Range["D27"].Value = "";
                        x.Range["D28"].Value = "";
                        x.Range["B27"].Value = "";
                        x.Range["B28"].Value = "";
                        x.Range["E27"].Value = "";
                        x.Range["E28"].Value = "";

                        x.Range["A30"].Value = "";
                        x.Range["D30"].Value = "";
                        x.Range["D31"].Value = "";
                        x.Range["B30"].Value = "";
                        x.Range["B31"].Value = "";
                        x.Range["E30"].Value = "";
                        x.Range["E31"].Value = "";

                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 2:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A21"].Value = "";
                        x.Range["D21"].Value = "";
                        x.Range["D22"].Value = "";
                        x.Range["B21"].Value = "";
                        x.Range["B22"].Value = "";
                        x.Range["E21"].Value = "";
                        x.Range["E22"].Value = "";

                        x.Range["A24"].Value = "";
                        x.Range["D24"].Value = "";
                        x.Range["D25"].Value = "";
                        x.Range["B24"].Value = "";
                        x.Range["B25"].Value = "";
                        x.Range["E24"].Value = "";
                        x.Range["E25"].Value = "";

                        x.Range["A27"].Value = "";
                        x.Range["D27"].Value = "";
                        x.Range["D28"].Value = "";
                        x.Range["B27"].Value = "";
                        x.Range["B28"].Value = "";
                        x.Range["E27"].Value = "";
                        x.Range["E28"].Value = "";

                        x.Range["A30"].Value = "";
                        x.Range["D30"].Value = "";
                        x.Range["D31"].Value = "";
                        x.Range["B30"].Value = "";
                        x.Range["B31"].Value = "";
                        x.Range["E30"].Value = "";
                        x.Range["E31"].Value = "";

                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 3:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A24"].Value = "";
                        x.Range["D24"].Value = "";
                        x.Range["D25"].Value = "";
                        x.Range["B24"].Value = "";
                        x.Range["B25"].Value = "";
                        x.Range["E24"].Value = "";
                        x.Range["E25"].Value = "";

                        x.Range["A27"].Value = "";
                        x.Range["D27"].Value = "";
                        x.Range["D28"].Value = "";
                        x.Range["B27"].Value = "";
                        x.Range["B28"].Value = "";
                        x.Range["E27"].Value = "";
                        x.Range["E28"].Value = "";

                        x.Range["A30"].Value = "";
                        x.Range["D30"].Value = "";
                        x.Range["D31"].Value = "";
                        x.Range["B30"].Value = "";
                        x.Range["B31"].Value = "";
                        x.Range["E30"].Value = "";
                        x.Range["E31"].Value = "";

                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 4:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A27"].Value = "";
                        x.Range["D27"].Value = "";
                        x.Range["D28"].Value = "";
                        x.Range["B27"].Value = "";
                        x.Range["B28"].Value = "";
                        x.Range["E27"].Value = "";
                        x.Range["E28"].Value = "";

                        x.Range["A30"].Value = "";
                        x.Range["D30"].Value = "";
                        x.Range["D31"].Value = "";
                        x.Range["B30"].Value = "";
                        x.Range["B31"].Value = "";
                        x.Range["E30"].Value = "";
                        x.Range["E31"].Value = "";

                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 5:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A30"].Value = "";
                        x.Range["D30"].Value = "";
                        x.Range["D31"].Value = "";
                        x.Range["B30"].Value = "";
                        x.Range["B31"].Value = "";
                        x.Range["E30"].Value = "";
                        x.Range["E31"].Value = "";

                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 6:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A33"].Value = "";
                        x.Range["D33"].Value = "";
                        x.Range["D34"].Value = "";
                        x.Range["B33"].Value = "";
                        x.Range["B34"].Value = "";
                        x.Range["E33"].Value = "";
                        x.Range["E34"].Value = "";

                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                    case 7:
                        #region Suppression lignes en trop
                        //Suppression des lignes en trop
                        x.Range["A36"].Value = "";
                        x.Range["D36"].Value = "";
                        x.Range["D37"].Value = "";
                        x.Range["B36"].Value = "";
                        x.Range["B37"].Value = "";
                        x.Range["E36"].Value = "";
                        x.Range["E37"].Value = "";
                        #endregion
                        break;
                }

                while (nbDate > 0)
                {
                    switch (nbDate)
                    {
                        case 1:
                            x.Range["A15"].Value = dtpDate1.Value.ToShortDateString();
                            x.Range["D15"].Value = txbxPlage1.Text;
                            x.Range["D16"].Value = txbxAccomp1.Text;

                            nbDate -= 1;
                            break;
                        case 2:
                            x.Range["A18"].Value = dtpDate2.Value.ToShortDateString();
                            x.Range["D18"].Value = txbxPlage2.Text;
                            x.Range["D19"].Value = txbxAccomp2.Text;

                            nbDate -= 1;
                            break;
                        case 3:
                            x.Range["A21"].Value = dtpDate3.Value.ToShortDateString();
                            x.Range["D21"].Value = txbxPlage3.Text;
                            x.Range["D22"].Value = txbxAccomp3.Text;

                            nbDate -= 1;
                            break;
                        case 4:
                            x.Range["A24"].Value = dtpDate4.Value.ToShortDateString();
                            x.Range["D24"].Value = txbxPlage4.Text;
                            x.Range["D25"].Value = txbxAccomp4.Text;

                            nbDate -= 1;
                            break;
                        case 5:
                            x.Range["A27"].Value = dtpDate5.Value.ToShortDateString();
                            x.Range["D27"].Value = txbxPlage5.Text;
                            x.Range["D28"].Value = txbxAccomp5.Text;

                            nbDate -= 1;
                            break;
                        case 6:
                            x.Range["A30"].Value = dtpDate6.Value.ToShortDateString();
                            x.Range["D30"].Value = txbxPlage6.Text;
                            x.Range["D31"].Value = txbxAccomp6.Text;
                            nbDate -= 1;
                            break;
                        case 7:
                            x.Range["A33"].Value = dtpDate7.Value.ToShortDateString();
                            x.Range["D33"].Value = txbxPlage7.Text;
                            x.Range["D34"].Value = txbxAccomp7.Text;
                            nbDate -= 1;
                            break;
                        case 8:
                            x.Range["A36"].Value = dtpDate8.Value.ToShortDateString();
                            x.Range["D36"].Value = txbxPlage8.Text;
                            x.Range["D37"].Value = txbxAccomp8.Text;
                            nbDate -= 1;
                            break;
                    }
                }

                string nomFichier = dirBeforeAppli + @"\Facture" + DateTime.Today.Year.ToString() + @"\" + "VEUILLEZ INDIQUER LA VILLE";

                Program.outils.getConnection().Open();
                requete5 = "Select [Ville] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
                cmd5.CommandText = requete5;
                dr5 = cmd5.ExecuteReader();

                while (dr5.Read())
                {
                    if (Program.outils.getNumFacture() < 100)
                    {
                        nomFichier = dirBeforeAppli + @"\Facture" + DateTime.Today.Year.ToString() + @"\5860" + (DateTime.Today.Year - 2000) + "-0" + Convert.ToString(Program.outils.getNumFacture() - 1) + " " + dr5[0].ToString();
                    }
                    else
                    {
                        nomFichier = dirBeforeAppli + @"\Facture" + DateTime.Today.Year.ToString() + @"\5860" + (DateTime.Today.Year - 2000) + "-" + Convert.ToString(Program.outils.getNumFacture() - 1) + " " + dr5[0].ToString();
                    }
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

                this.Hide();
                Facture facture = new Facture();
                facture.Show();
            }

        }

        private void cbxVille_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Remplissage txbxAdresse
            Program.outils.getConnection().Open();
            string requete = "Select [LigneAdr1] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete1 = "Select [LigneAdr2] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete2 = "Select [LigneAdr3] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete3 = "Select [LigneAdr4] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete4 = "Select [CodePostal] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";
            string requete5 = "Select [Ville] from FactureAdresse where [Ville] = \"" + cbxVille.Text + "\";";

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

        private void btAddDate_Click(object sender, EventArgs e)
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
                    grbDate7.Visible = true;
                    nbDate += 1;
                    break;
                case 7:
                    grbDate8.Visible = true;
                    nbDate += 1;
                    break;
                case 8:
                    MessageBox.Show("Pour le moment vous êtes limiter à 8 dates par facture, une des solutions serait de faire la facture manuellement.", "Limite de Date :", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
            }
        }

        private void btSuppDate_Click(object sender, EventArgs e)
        {
            switch (nbDate)
            {
                case 1:
                    MessageBox.Show("Afin d'établir la facture vous devez au moins renseigner une date.", "Date minimum :", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                case 7:
                    grbDate7.Visible = false;
                    nbDate -= 1;
                    break;
                case 8:
                    grbDate8.Visible = false;
                    nbDate -= 1;
                    break;
            }
        }

        private void btAddVille_Click_1(object sender, EventArgs e)
        {
            AjoutAdresseFacture addAdresse = new AjoutAdresseFacture();
            addAdresse.Show();
        }
    }
}
