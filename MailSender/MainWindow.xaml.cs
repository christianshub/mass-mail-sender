using System;
using System.Windows;
using System.Net.Mail;
using System.Net;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Threading;

namespace MailSender
{    


    public partial class MainWindow : System.Windows.Window
    {
 
        public List<string> LstNames;
        public List<string> LstMails;

        public MainWindow()
        {
            InitializeComponent();
        }


        /// <summary>
        ///    Sends the mails
        /// </summary>
        private void BtnSend_Click(object sender, RoutedEventArgs e)
        {

            //var smtpServerName = ConfigurationManager.AppSettings["SmtpServer"];
            var smtpServerName = "smtp.gmail.com";
            var port = "587";
            //var port = ConfigurationManager.AppSettings["Port"];

            // var senderEmailId = ConfigurationManager.AppSettings["SenderEmailId"];
            // var senderPassword = ConfigurationManager.AppSettings["SenderPassword"];

            var smptClient = new SmtpClient(smtpServerName, Convert.ToInt32(port))
            {
                Credentials = new NetworkCredential(loginBox.Text, passwordBox.Password.ToString()),
                EnableSsl = true
            };

            if (LstNames.Count == LstMails.Count && LstNames.Count > 0)
            {
                for (var i = 0; i < LstNames.Count; i++)
                {
                    var fullText = "Kære " + LstNames[i] + "," + "\n\n" + Indhold.Text;
                    smptClient.Send(loginBox.Text, LstMails[i], Emne.Text, fullText);
                    //MessageBox.Show(fullText);
                    //MessageBox.Show("Login: " + loginBox.Text + " Password: " + passwordBox.Password.ToString());
                    Thread.Sleep(1000);
                }
            }
            else
            {
                MessageBox.Show("Something's wrong, LstNames.Count:" + LstNames.Count);
            }

            Emne.Text = "";
            Indhold.Text = "";
            Emne.Focus();
        }


        /// <summary>
        ///    Reset fields
        /// </summary>
        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            Emne.Text = "";
            Indhold.Text = "";
            Emne.Focus();
        }     

        /// <summary>
        ///    Open File -> Return two lists of string names (list 1: Name, list 2: Email)
        /// </summary>
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
   

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls|Excel Files(*.xlsm)|*.xlsm";

            var result = openFileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                // Open document 
                string filePath = openFileDialog.FileName;

                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                
                var xlWorkBook  = xlApp.Workbooks.Open(@filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                var xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.get_Item(1);
                var range = xlWorkSheet.UsedRange;


                int count = 0;

                for (int i = 1; i < 5; i++)
                {
                    var name = (string) xlWorkSheet.Cells[i, 1].Value;  // row 1, cell i

                    if (name != null)
                    {
                        count++;
                    }
                }

                LstNames = new List<string>();
                LstMails = new List<string>();

                for (int i = 1; i <= count; i++)
                {

                    var name = (string) xlWorkSheet.Cells[i, 1].Value;  // row 1, cell i
                    var mail = (string) xlWorkSheet.Cells[i, 2].Value;  // row 1, cell i

                    LstNames.Add(name);
                    LstMails.Add(mail);
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                textBox.Text = filePath;
            }
        }
    }
}