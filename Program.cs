using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

using System.Linq;
using System.Text;

using System.IO;
using System.Diagnostics;

using System.Runtime.InteropServices;
using System.Threading;

namespace ExplorerTracker
{
    class Program
    {
        public static string time,npad,li,line;
        public static string timet = DateTime.Now.ToString("MM-dd-yy");

        public static int tt = 0,fl,chk=0;
        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lp1, string lp2);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static void copyDirectory(string strSource, string strDestination)
        {
            if (!Directory.Exists(strDestination))
            {
                Directory.CreateDirectory(strDestination);
            }

            DirectoryInfo dirInfo = new DirectoryInfo(strSource);
            FileInfo[] files = dirInfo.GetFiles();
            foreach (FileInfo tempfile in files)
            {
                try
                {
                    tempfile.CopyTo(Path.Combine(strDestination, tempfile.Name));
                }
                catch
                { }
            }

            DirectoryInfo[] directories = dirInfo.GetDirectories();
            foreach (DirectoryInfo tempdir in directories)
            {
                try
                {
                    copyDirectory(Path.Combine(strSource, tempdir.Name), Path.Combine(strDestination, tempdir.Name));
                }
                catch { }
            }

        }

        static void Main(string[] args)
        {
            
               
                String[] save = new String[10000];
                int last = 0;

                int dont, i, l, ii, j, ct;
                IntPtr MyHwnd;

                String dt = "";
                string dtemp = "";
                string strSource2, strDestination2, pth;
                strSource2 = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
                strDestination2 = Environment.CurrentDirectory + "\\rct";

                copyDirectory(strSource2, strDestination2);
                //MessageBox.Show("succesfull");
                string c = "";
                pth = strDestination2 + "\\rct " + timet + c + ".txt";// +DateTime.Now.ToString() + ".txt";
                int y = 0;

                while (File.Exists(pth))
                {
                    c = y.ToString(); y++;
                    pth = strDestination2 + "\\rct " + timet + " ( " + c + " )" + ".txt";
                }
                try
                {
                    File.Create(pth).Close();

                    TextWriter tw = new StreamWriter(pth);

                    tw.Close();
                }
                catch { }

                while (true)
                {

                    using (StreamWriter w = File.AppendText(pth))
                    {

                        if (tt %450==0)
                        {
                            chk = 0;
                            time = "\n\n"+DateTime.Now.ToString() + "\n";
                            w.WriteLine(time);
                            w.WriteLine("");
                        }


                        w.Close();
                    }

                    if (tt >= 2*1800)
                    {
                        string lc = Environment.CurrentDirectory;
                        try
                        {
                            System.Diagnostics.Process.Start(@lc+"\\ExplorerTracker.exe");
                            Environment.Exit(1);
                            //Application.Exit();
                        }
                        catch { }
                    
                    }


                    MyHwnd = FindWindow(null, "Directory");
                    var t = Type.GetTypeFromProgID("Shell.Application");
                    dynamic o = Activator.CreateInstance(t);
                    try
                    {
                        var ws = o.Windows();
                        for (i = 0; i < ws.Count; i++)
                        {

                            var ie = ws.Item(i);

                            // if (ie == null || ie.hwnd != (long)MyHwnd) continue;
                            var path = System.IO.Path.GetFileName((string)ie.FullName);
                            //dt += Convert.ToString(path);
                            try
                            {
                                if (path.ToLower() == "explorer.exe")
                                {
                                    var explorepath = ie.document.focuseditem.path;
                                    dtemp = Convert.ToString(explorepath);
                                    l = dtemp.Length;
                                    while (dtemp[l - 1] != '\\')
                                    {

                                        l--;
                                    }
                                    for (ii = 0; ii < l; ii++)
                                        dt += dtemp[ii];
                                    dt += "\n";
                                    dont = 0;
                                    for (j = last, ct = 0; ct <= chk && j >= 0; j--, ct++)
                                    {
                                        if (dt == save[j])
                                            dont = 1;


                                    }
                                    
                                    //

                                /*    using (System.IO.StreamReader file = new System.IO.StreamReader(@pth))
                                    {

                                        
                                         fl = 0;
                                        j=0;
                                        li = file.ReadLine();
                                        if (new FileInfo(pth).Length == 0)
                                            fl = 1;
                                        while (fl == 0)
                                        {j++;
                                            line += li;
                                            li = file.ReadLine();
                                            if (li == null)
                                                fl = 1;

                                            if (j > last - (ws.Count * 2) && j >= 0)
                                            {
                                                if (dt ==li)
                                                    dont = 1;
                                            
                                            }

                                        }
                                        
                                    }*/
                                    //


                                    if (dont == 0)
                                    {
                                        save[last] = dt; 
                                        last++;
                                        chk++;
                                        using (StreamWriter w = File.AppendText(pth))
                                        {

                                            w.WriteLine(dt);
                                            w.Close();
                                        }

                                    }


                                    dtemp = "";
                                    dt = "";
                                }
                            }
                            catch { }
                        }
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(o);

                    }




                    Thread.Sleep(2000);
                    tt++;
                    
                }
            
        }
    }
}
