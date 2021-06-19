using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using Microsoft.Win32;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.IO;
using OfficeOpenXml;
using System.Threading;
using System.Windows.Threading;

namespace AutoMail.ViewModel
{
    public class MainViewModel : BaseViewModel
    {
        // https://www.google.com/settings/u/1/security/lesssecureapps

        public ICommand AttachCommand { get; set; }
        public ICommand SentMailCommand { get; set; }
        public ICommand PasswordChangedCommand { get; set; }
        public ICommand TextBoxChangedCommand { get; set; }
        public ICommand CheckedCommand { get; set; }
        public ICommand UncheckedCommand { get; set; }
        public ICommand BrowserExcelCommand { get; set; }
        public string Username { get => username; set { username = value; OnPropertyChanged(); } }
        public string Password { get => password; set { password = value; OnPropertyChanged(); } }
        public string ToAddress { get => toAddress; set { toAddress = value; OnPropertyChanged(); } }
        public string Subject { get => subject; set { subject = value; OnPropertyChanged(); } }
        public string Message { get => message; set { message = value; OnPropertyChanged(); } }
        public string FileAttach { get => fileAttach; set { fileAttach = value; OnPropertyChanged(); } }
        public bool ModeSent { get => modeSent; set { modeSent = value; OnPropertyChanged(); } }
        public string ColNumStr { get => colNumStr; set { colNumStr = value; OnPropertyChanged(); } }
        public string SheetNumStr { get => sheetNumStr; set { sheetNumStr = value; OnPropertyChanged(); } }

        private string username;
        private string password;
        private string toAddress;
        private string subject;
        private string message;
        private string fileAttach;
        private string colNumStr;
        private string sheetNumStr;
        private Attachment attachment;
        private bool modeSent;
        private int colNum;
        private int sheetNum;

        static object syncObj = new object();
        static int countSuccess = 0;

        public MainViewModel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ModeSent = true;

            AttachCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
                {
                    OpenFileDialog dialog = new OpenFileDialog();
                    if (dialog.ShowDialog() == true)
                        FileAttach = dialog.FileName;
                });

            SentMailCommand = new RelayCommand<object>((p) => 
                {
                    if (string.IsNullOrEmpty(Username) || string.IsNullOrEmpty(Password) || string.IsNullOrEmpty(ToAddress))
                        return false;
                    return true; 
                },
                (p) =>
                {
                    attachment = null;
                    try
                    {
                        FileInfo file = new FileInfo(FileAttach);
                        attachment = new Attachment(file.FullName);
                    }
                    catch
                    {

                    }

                    List<string> listMail;
                    if (ModeSent == false)
                    {
                        bool res1 = ConvertToInt(ColNumStr, out colNum);
                        bool res2 = ConvertToInt(SheetNumStr, out sheetNum);
                        if (!res1 || !res2)
                        {
                            MessageBox.Show("The Sheet and Column must be a unsigned numeric");
                            return;
                        }
                        listMail = GetListMailFromXLSX(sheetNum, colNum, ToAddress);
                        List<string> errorMail = new List<string>();
                        foreach (string item in listMail)
                        {
                            Thread thread = new Thread(() =>
                            {
                                if (!SentMail(item, attachment))
                                    errorMail.Add(item);
                                else
                                {
                                    lock (syncObj)
                                    {
                                        countSuccess++;
                                    }
                                }
                            });
                            thread.Start();
                        }
                        string log = "";
                        foreach (string item in errorMail)
                        {
                            log += item + "\n";
                        }
                        MessageBox.Show(string.Format("Send {0}/{1} successfully.\nList error mail:\n{2}", countSuccess, listMail.Count, log));
                    }
                    else
                    {
                        Thread thread = new Thread(() =>
                        {
                            if (!SentMail(ToAddress, attachment))
                                MessageBox.Show("Send to "+ ToAddress + " is failed.");
                            else
                                MessageBox.Show("Success");
                        });
                        thread.Start();
                    }
                }
                );

            PasswordChangedCommand = new RelayCommand<PasswordBox>((p) => { return true; }, (p) =>
                {
                    Password = p.Password;
                });

            TextBoxChangedCommand = new RelayCommand<System.Windows.Controls.TextBox>((p) => 
            {
                return true; 
            },
            (p) =>
            {
                if (p.Name == "txbUsername")
                    Username = p.Text;
                else if (p.Name == "txbSubject")
                    Subject = p.Text;
                else if (p.Name == "txbTo")
                    ToAddress = p.Text;
                else if (p.Name == "txbColNum")
                    ColNumStr = p.Text;
                else if (p.Name == "txbSheet")
                    SheetNumStr = p.Text;
            });

            CheckedCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                ModeSent = true;
            });

            UncheckedCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                ModeSent = false;
            });

            BrowserExcelCommand = new RelayCommand<RadioButton>((p) => 
            {
                if (ModeSent) return false;
                return true;
            }, 
            (p) =>
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel | *.xlsx";
                if (dialog.ShowDialog() == true)
                    ToAddress = dialog.FileName;
            });
        }

        bool SentMail(string toMail, Attachment attachment = null)
        {
            MailMessage mailMessage = new MailMessage(Username, toMail, Subject, Message);

            if (attachment != null)
            {
                mailMessage.Attachments.Add(attachment);
            }

            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new NetworkCredential(Username, Password);
            try
            {
                smtpClient.Send(mailMessage);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        bool ConvertToInt(string str, out int n)
        {
            bool res = Int32.TryParse(str, out n);
            if (res)
            {
                if (n >= 1)
                {
                    return true;
                }
                else
                    MessageBox.Show("The Column and Sheet must be a unsigned numeric.");
            }
            return false;
        }

        List<string> GetListMailFromXLSX(int sheet, int col, string file)
        {
            //List mail
            List<string> listMail = new List<string>();

            //open file
            var package = new ExcelPackage(new FileInfo(file));

            //open sheet
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheet-1];

            //traverse in column col
            for(int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
            {
                try
                {
                    var temp = workSheet.Cells[i, col].Value;
                    if (temp != null)
                    {
                        string mail = temp.ToString();
                        if (!string.IsNullOrEmpty(mail) && mail.Contains("@"))
                            listMail.Add(mail);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error");
                }
            }
            return listMail;
        }
    }
}
