using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WindowsForms1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string test = String.Empty;
            test = "Test String";
            test = "Test\"\"\n\t";
            test = @"Test String";
            test = @"Test""String""";
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}

namespace WindowsForms1.Test
{
    static class Program
    {
        static void Test()
        {
            string test = String.Empty;
            test = "Test String";
        }
    }
}