namespace DevisTorcy
{
    partial class Menu
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.devisToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.factureToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dPAEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.devisToolStripMenuItem,
            this.factureToolStripMenuItem,
            this.dPAEToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // devisToolStripMenuItem
            // 
            this.devisToolStripMenuItem.Name = "devisToolStripMenuItem";
            this.devisToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.devisToolStripMenuItem.Text = "Devis";
            this.devisToolStripMenuItem.Click += new System.EventHandler(this.devisToolStripMenuItem_Click);
            // 
            // factureToolStripMenuItem
            // 
            this.factureToolStripMenuItem.Name = "factureToolStripMenuItem";
            this.factureToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.factureToolStripMenuItem.Text = "Facture";
            this.factureToolStripMenuItem.Click += new System.EventHandler(this.factureToolStripMenuItem_Click);
            // 
            // dPAEToolStripMenuItem
            // 
            this.dPAEToolStripMenuItem.Name = "dPAEToolStripMenuItem";
            this.dPAEToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.dPAEToolStripMenuItem.Text = "DPAE";
            this.dPAEToolStripMenuItem.Click += new System.EventHandler(this.dPAEToolStripMenuItem_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
            this.pictureBox1.Location = new System.Drawing.Point(193, 144);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(400, 147);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Menu";
            this.Text = "Menu";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem devisToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem factureToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dPAEToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

