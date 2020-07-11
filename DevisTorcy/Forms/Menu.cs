using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DevisTorcy
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {   

        }

        private void devisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Devis devis = new Devis();
            devis.Show();
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

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
