namespace DevisTorcy
{
    partial class Facture
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Facture));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.factureToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dPAEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem,
            this.factureToolStripMenuItem,
            this.dPAEToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 2;
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.menuToolStripMenuItem.Text = "Devis";
            this.menuToolStripMenuItem.Click += new System.EventHandler(this.menuToolStripMenuItem_Click);
            // 
            // factureToolStripMenuItem
            // 
            this.factureToolStripMenuItem.Name = "factureToolStripMenuItem";
            this.factureToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.factureToolStripMenuItem.Text = "Menu";
            this.factureToolStripMenuItem.Click += new System.EventHandler(this.factureToolStripMenuItem_Click);
            // 
            // dPAEToolStripMenuItem
            // 
            this.dPAEToolStripMenuItem.Name = "dPAEToolStripMenuItem";
            this.dPAEToolStripMenuItem.Size = new System.Drawing.Size(47, 20);
            this.dPAEToolStripMenuItem.Text = "DPAE";
            this.dPAEToolStripMenuItem.Click += new System.EventHandler(this.dPAEToolStripMenuItem_Click);
            // 
            // Facture
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Facture";
            this.Text = "Facture";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem factureToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dPAEToolStripMenuItem;
    }
}