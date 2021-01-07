using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace universal_logger
{
    static class Program
    {
        /// <summary>Az alkalmazás fő belépési pontja</summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainWindow());
        }
    }
}
