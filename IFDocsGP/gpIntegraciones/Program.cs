using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace gp.InterfacesPersonalizadas
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] argumento)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new winformGeneraFE());

        }
    }
}
