using System;
using System.Windows.Forms;
using System.Data.OleDb;

namespace DevisTorcy.Forms
{

    public partial class AjoutAdresseDevis : Form
    {

        public AjoutAdresseDevis()
        {
            InitializeComponent();
        }

        private void btSend_Click(object sender, EventArgs e)
        {
            Program.outils.getConnection().Open();
            string requete = "Insert into DevisAdresse VALUES (" + txbxCP.Text + ",\"" + txbxNomVille.Text + "\",\"" + txbxLigne1.Text + "\",\"" + txbxLigne2.Text + "\",\"" + txbxLigne3.Text + "\",\"" + txbxLigne4.Text + "\");";
            OleDbCommand cmd = new OleDbCommand(requete, Program.outils.getConnection());
            cmd.CommandText = requete;
            cmd.ExecuteNonQuery();
            Program.outils.getConnection().Close();
            this.Hide();
        }
    }
}
