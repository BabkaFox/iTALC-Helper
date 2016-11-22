using System;
using System.Collections.Generic;
using System.Windows;
using System.DirectoryServices;


namespace iTALCClassHelper
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string FILE_SETTINGS = @"C:\ProgramData\iTALC\config\GlobalConfig.xml";
        private static String curCompName = Environment.MachineName;

        public MainWindow()
        {
            InitializeComponent();

            List<string> ComputerNames = new List<string>();

            DirectoryEntry entry = new DirectoryEntry("LDAP://OU=Moscow,DC=class,DC=mfua,DC=ru");
            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = ("(objectClass=computer)");
            mySearcher.SizeLimit = int.MaxValue;
            mySearcher.PageSize = int.MaxValue;

            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"^[a-zA-Z]+-\d{1,3}-\d{1,4}-\d{1,4}\b$");

            Char delimiter = '-';

            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                string ComputerName = resEnt.GetDirectoryEntry().Name;

                if (ComputerName.StartsWith("CN=" + curCompName.Split(delimiter)[0] + "-" + curCompName.Split(delimiter)[1] + "-" + curCompName.Split(delimiter)[2]))
                    ComputerName = ComputerName.Remove(0, "CN=".Length);
                if (reg.IsMatch(ComputerName))
                    ComputerNames.Add(ComputerName);
            }

            mySearcher.Dispose();
            entry.Dispose();

            using (System.IO.StreamWriter sw = System.IO.File.CreateText(FILE_SETTINGS))
            {
                sw.WriteLine("<?xml version=\"1.0\"?>");
                sw.WriteLine("<!DOCTYPE italc-config-file>");
                sw.WriteLine("<globalclientconfig version=\"3.0.1\">");
                sw.WriteLine("<body>");
                sw.WriteLine("<classroom name=\"" + curCompName.Split(delimiter)[0] + "-" + curCompName.Split(delimiter)[1] + "-" + curCompName.Split(delimiter)[2] + "\">");

                foreach (var item in ComputerNames)
                {
                    sw.WriteLine("<client mac=\"\" name=\"\" hostname=\"" + item + "\" id=\"\" type=\"1\"/>");
                }

                sw.WriteLine("</classroom>");
                sw.WriteLine("</body>");
                sw.WriteLine("</globalclientconfig>");
            }
            return;
            
        }
    }
}
