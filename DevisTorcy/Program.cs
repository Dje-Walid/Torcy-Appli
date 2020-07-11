using System;
using System.Windows.Forms;

namespace DevisTorcy
{
    static class Program
    {
        public static Outils outils = new Outils();
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Menu());
        }
    }

}
