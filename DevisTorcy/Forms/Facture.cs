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
    public partial class Facture : Form
    {
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
    }
}
