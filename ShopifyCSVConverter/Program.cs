using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    static class Program
    {
        static Mutex _m;

        static bool InstanceExists()
        {            
            if(!Mutex.TryOpenExisting("SHOPIFY_CSV_CONVERTER", out _m))
            {
                _m = new Mutex(true, "SHOPIFY_CSV_CONVERTER");
                return false;
            }
            return true;
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (InstanceExists()) return;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Converter());
        }
    }
}
