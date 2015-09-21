using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace EmailExtractor
{
    static class l5
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new l4());
        }
    }
}
