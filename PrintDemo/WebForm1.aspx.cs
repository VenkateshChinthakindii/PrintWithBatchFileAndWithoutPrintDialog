using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Management;
using System.Drawing.Printing;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using System.Reflection;
//using System.Threading;
using SHDocVw;
using Microsoft.Win32;
using System.Data.SqlClient;

namespace PrintDemo
{
    public partial class WebForm1 : System.Web.UI.Page
    {

        private bool documentLoaded;
        private bool documentPrinted;

        private void ie_DocumentComplete(object pDisp, ref object URL)
        {
            documentLoaded = true;
        }

        private void ie_PrintTemplateTeardown(object pDisp)
        {
            documentPrinted = true;
        }

        public void Print(string htmlFilename)
        {
            string strKey = "Software\\Microsoft\\Internet Explorer\\PageSetup";
            bool bolWritable = true;
            string strName = "footer";
            object oValue = "";
            RegistryKey oKey = Registry.CurrentUser.OpenSubKey(strKey, bolWritable);
            Console.Write(strKey);
            oKey.SetValue(strName, oValue);
            oKey.Close();

            documentLoaded = false;
            documentPrinted = false;

            InternetExplorer ie = new InternetExplorerClass();
            ie.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(ie_DocumentComplete);
            ie.PrintTemplateTeardown += new DWebBrowserEvents2_PrintTemplateTeardownEventHandler(ie_PrintTemplateTeardown);

            object missing = Missing.Value;

            ie.Navigate(htmlFilename, ref missing, ref missing, ref missing, ref missing);
            while (!documentLoaded && ie.QueryStatusWB(OLECMDID.OLECMDID_PRINT) != OLECMDF.OLECMDF_ENABLED)
                Thread.Sleep(100);

            ie.ExecWB(OLECMDID.OLECMDID_PRINT, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref missing, ref missing);
            while (!documentPrinted)
                Thread.Sleep(100);

            ie.DocumentComplete -= ie_DocumentComplete;
            ie.PrintTemplateTeardown -= ie_PrintTemplateTeardown;
            ie.Quit();
        }


        public void getAttachment()
        {
            //SqlConnection con = new SqlConnection();
            string constr = "server=WXSQL003;database=Finance;user=FinanceUser;password=FinanceUser";
            //con.ConnectionString = constr;
            using (SqlConnection connection = new SqlConnection(constr))
            using (SqlCommand cmd = new SqlCommand("select * from [Finance].[dbo].[Payment_RequestAttachment] where id=13", connection))
            {
                connection.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    // Check is the reader has any rows at all before starting to read.
                    if (reader.HasRows)
                    {
                        // Read advances to the next row.
                        while (reader.Read())
                        {

                            string FileName = reader.GetString(reader.GetOrdinal("FileName"));
                            int attachment = reader.GetOrdinal("Attachment");
                            byte[] attachmentByte = new byte[963107];
                            // If a column is nullable always check for DBNull...
                            if (!reader.IsDBNull(attachment))
                            {

                                reader.GetBytes(attachment, 0, attachmentByte, 0, 963107);
                            }
                            //using (var myFile = File.Create(@"E:\Practice\PrintDemo\PrintDemo\" + FileName))
                            //{
                            //    // interact with myFile here, it will be disposed automatically
                            //}                           

                            File.WriteAllBytes(@"E:\Practice\PrintDemo\PrintDemo\" + FileName, attachmentByte); // Requires System.IO
                        }
                    }
                }
            }
        }
        /// <summary>
        /// ///////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void myProcess_Exited(object sender, System.EventArgs e)
        {
            //eventHandled = true;

            Console.WriteLine("print executed");
        }

        //WebBrowser webBrowser = null;
        //void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        //{
        //    webBrowser.Print();
        //}
        //public void StartThred()
        //{
        //    webBrowser = new WebBrowser();
        //    webBrowser.DocumentCompleted += webBrowser_DocumentCompleted;
        //    webBrowser.DocumentText = "<html><body>..some html code..</body></html>";
        //    webBrowser.Print();
        //}
        protected void Page_Load(object sender, EventArgs e)
        {
            //ProcessStartInfo info = new ProcessStartInfo();
            //info.Verb = "print";            
            //info.FileName = @"E:\Practice\PrintFilesDi\Recent changes sp.docx";
            //info.CreateNoWindow = true;
            //info.WindowStyle = ProcessWindowStyle.Hidden;

            //Process p = new Process();
            //p.StartInfo = info;
            //p.Start();

            //getAttachment();
            PrinterSettings pSett = new PrinterSettings();
            if (pSett.PrinterName != null)
            {
                foreach (string printerName in PrinterSettings.InstalledPrinters)
                {
                    if (printerName.ToLower() == "\\\\inhydprintsrv01\\printerzone_1")
                    {
                        Process myProcess = new Process();
                        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        myProcess.StartInfo.FileName = Server.MapPath("~/batchFile.bat");
                        myProcess.StartInfo.Arguments = String.Format("{0} {1} {2}", @"\\inhydprintsrv01\printerzone_1", @"E:\Practice\PrintFilesDi", @"E:\Practice\PrintDemo\PrintDemo\printjs.bat");
                        //myProcess.StartInfo.Verb = "Print";
                        myProcess.StartInfo.CreateNoWindow = true;
                        myProcess.StartInfo.UseShellExecute = false;
                        myProcess.EnableRaisingEvents = true;
                        myProcess.Exited += new EventHandler(myProcess_Exited);
                        myProcess.Start();
                        //myProcess.WaitForExit();
                        //int a = myProcess.ExitCode;
                    }
                }
            }

            //To print html pages without dialog box
            // Print(HttpContext.Current.Server.MapPath("~/HtmlPage1.html"));

            #region Old code commented
            //FileStream fs = null;
            //if (!File.Exists(Server.MapPath("~/batchFile.bat")))
            //{
            //    using (fs = File.Create(Server.MapPath("~/batchFile.bat")))
            //    {

            //    }
            //    using (StreamWriter sw = new StreamWriter(Server.MapPath("~/batchFile.bat")))
            //    {
            //        sw.WriteLine("print "+ HttpContext.Current.Server.MapPath("~/AFE Find Default Behavior.docx") + " /c /d:\\inhydprintsrv01\\PrinterZone_1");
            //        //sw.WriteLine("SET file="+ HttpContext.Current.Server.MapPath("~/AFE Find Default Behavior.docx"));
            //        //sw.WriteLine("RUNDLL32.EXE PRINTUI.DLL,PrintUIEntry /Xg /n \" % PrinterName % \" /f \" % file % \" /q");
            //        //sw.WriteLine("IF EXIST \" % file % \" (");
            //        //sw.WriteLine("SET PrinterName=\\inhydprintsrv01\\PrinterZone_1");
            //    }
            //}
            //Process pro = new Process();
            //ProcessStartInfo start = new ProcessStartInfo();

            //start.WindowStyle = ProcessWindowStyle.Hidden;
            //start.Arguments = string.Format("/p /h {0}", HttpContext.Current.Server.MapPath("~/AFE Find Default Behavior.docx"));  //"\\inhydprintsrv01\\PrinterZone_1";

            //start.FileName = HttpContext.Current.Server.MapPath("~/AFE Find Default Behavior.docx");
            //start.Verb = "print";
            ////[DefaultValue(ProcessWindowStyle.Normal)]
            //start.WindowStyle = ProcessWindowStyle.Hidden;
            //start.CreateNoWindow = true; start.UseShellExecute = true;
            //pro.StartInfo = start;

            //pro.EnableRaisingEvents = true;
            //pro.Start();
            //pro.WaitForInputIdle();
            //System.Threading.Thread.Sleep(3000);
            //if (false == pro.CloseMainWindow())
            //    pro.Kill();
            //PrintFiles("\\YPHprinterPh3DT.Yash.com\\HP LaserJet Pro MFP M127-M128 PCLmS", HttpContext.Current.Server.MapPath("~/HtmlPage1.html"), HttpContext.Current.Server.MapPath("~/HtmlPage1.html"));
            // PrintFiles("\\inhydprintsrv01\\PrinterZone_1", HttpContext.Current.Server.MapPath("~/AFE Find Default Behavior.docx"),
            // HttpContext.Current.Server.MapPath("~/HtmlPage1.html"));

            //string printContent = "sssssss\nssssss";
            //PrintFiles("\\YPHprinterPh3DT.Yash.com\\HP Laser Jet Pro MFP M127-M128 PCLmS",
            //    "AFE Find Default Behavior.docx"); 
            #endregion


        }
        #region Commneted Code
        //public void PrintFiles(string printerName, params string[] fileNames)
        //{
        //    //var files = String.Join(" ", fileNames);

        //    //var command = String.Format("/C print /D:{0} {1}", printerName, files);
        //    //ProcessStartInfo psInfo = new ProcessStartInfo();
        //    //psInfo.Arguments = String.Format(" -dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies=1 -sDEVICE=ljet4 -sOutputFile=\"\\\\spool\\{0}\" \"{1}\"", printerName, HttpContext.Current.Server.MapPath("~/HtmlPage1.html"));
        //    //psInfo.FileName = @"C:\Program Files\gs\gs8.70\bin\gswin32c.exe";
        //    //psInfo.UseShellExecute = false;
        //    //Process process = Process.Start(psInfo);

        //    //var files = String.Join(" ", fileNames);

        //    //var command = String.Format("/C print /D:{0} {1}", files, printerName);
        //    //var process = new Process();
        //    //var startInfo = new ProcessStartInfo
        //    //{
        //    //    WindowStyle = ProcessWindowStyle.Hidden,
        //    //    FileName = "cmd.exe",
        //    //    Arguments = command
        //    //};

        //    //process.StartInfo = startInfo;
        //    //process.Start();

        //    //Process printJob = new Process();
        //    //printJob.StartInfo.FileName = fileNames[0];
        //    //printJob.StartInfo.UseShellExecute = true;
        //    //printJob.StartInfo.Verb = "printto";
        //    //printJob.StartInfo.CreateNoWindow = true;
        //    //printJob.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //    //printJob.StartInfo.Arguments = string.Format("{0} {1}", printerName, fileNames[0]);
        //    ////printJob.StartInfo.Arguments = "\"" + printerAddress + "\"" + " " + "Microsoft XPS Document Writer";
        //    //printJob.StartInfo.WorkingDirectory =
        //    //    Path.GetDirectoryName(fileNames[0]);
        //    //printJob.Start();
        //    //System.IO.File.Copy(@"E:\Practice\PrintDemo\PrintDemo\AFE Find Default Behavior.docx", 
        //    //    printerName); ;
        //    //setDefaultPrinter(printerName);
        //    //ProcessStartInfo info = new ProcessStartInfo();
        //    //info.Verb = "print";
        //    //info.FileName = @"E:\Practice\PrintDemo\PrintDemo\AFE Find Default Behavior.docx";
        //    //info.CreateNoWindow = true;
        //    //info.WindowStyle = ProcessWindowStyle.Hidden;
        //    //info.UseShellExecute = true;
        //    //info.Arguments=printerName+" "+ @"E:\Practice\PrintDemo\PrintDemo\AFE Find Default Behavior.docx";
        //    //Process p = new Process();
        //    //p.StartInfo = info;
        //    //p.Start();

        //    //long ticks = -1;
        //    //while (ticks != p.TotalProcessorTime.Ticks)
        //    //{
        //    //    ticks = p.TotalProcessorTime.Ticks;
        //    //    Thread.Sleep(0000);
        //    //}

        //    System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
        //    info.Verb = "print";
        //    System.Diagnostics.Process p = new System.Diagnostics.Process();
        //    //Load Files in Selected Folder
        //    ///string[] allFiles = System.IO.Directory.GetFiles(Directory);
        //    foreach (string file in fileNames)
        //    {
        //        info.FileName = @file;
        //        info.CreateNoWindow = true;
        //        info.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
        //        p.StartInfo = info;
        //        p.Start();
        //        p.Refresh();
        //    }
        //    //p.Kill(); Can Create A Kill Statement Here... but I found I don't need one
        //    //MessageBox.Show("Print Complete");

        //    //var pi = new ProcessStartInfo(fileNames[0]);
        //    //pi.UseShellExecute = true;
        //    //pi.Verb = "print";
        //    //var process = System.Diagnostics.Process.Start(pi);
        //    //if (false == p.CloseMainWindow())
        //    //    p.Kill();

        //}
        //[DllImport("winspool.drv",CharSet=CharSet.Auto,SetLastError =true)]
        //public static extern bool SetDefaultPrinter(string printerName);
        //public void setDefaultPrinter(string printerName)
        //{

        //    SetDefaultPrinter(printerName);
        //} 
        #endregion
    }

}