using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Xml.Linq;
using static MailingWPF.MainWindow;
using SevenZip;

namespace MailingWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    public partial class MainWindow : Window
    {
        public ObservableCollection<ChosenFile> FileList { get; set; }

        public MainWindow()
        {
            FileList = new ObservableCollection<ChosenFile>();
            InitializeComponent();
            this.DataContext = this;

            //FileMessage.Text = Path.Combine(AppContext.BaseDirectory, "7z.dll") + Path.Combine(AppContext.BaseDirectory, "7z.dll") + Path.Combine(AppContext.BaseDirectory, "7z.dll") + Path.Combine(AppContext.BaseDirectory, "7z.dll");

            string[] args = Environment.GetCommandLineArgs();
            foreach (string s in args)
            {
                if (!Path.GetFileName(s).EndsWith(".dll"))
                {
                    string fileName = ReduceName(s, FileList);
                    FileList.Add(new ChosenFile(s, fileName));
                }
            }
        }

        public class ChosenFile
        {
            public ChosenFile(string Path, string Name)
            {
                filePath = Path;
                fileName = Name;
            }
            public string filePath { get; set; }
            public string fileName { get; set; }
        }

        private void dropfiles(object sender, DragEventArgs e)
        {
            string[] droppedFiles = null;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                droppedFiles = e.Data.GetData(DataFormats.FileDrop, true) as string[];
            }

            if ((null == droppedFiles) || (!droppedFiles.Any())) { return; }

            foreach (string s in droppedFiles)
            {
                string fileName = ReduceName(s, FileList);

                FileList.Add(new ChosenFile(s, fileName));
            }
        }

        //Interface
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            WindowState = WindowState.Minimized;

            string zipFile = "";

            //ZIP
            try
            {
                zipFile = CreateZipFolder(FileList);
            }
            catch(Exception exep)
            {
                FileMessage.Text = $"Zip Failed: {exep}";
                WindowState = WindowState.Normal;
                return;
            }


            if (zipFile == "toobig")
            {
                FileMessage.Text = "The Zip File is larger than 25mb";
                WindowState = WindowState.Normal;
                return;
            }

            //MAIL SUBJECT AND BODY
            string MailBody = "";
            string MailSubject = "";
            string permanentMessage = "If any file is denoted with 'HR' in the filename, it is intended for High-Resolution print.";
            int count = 0;

            MailBody += "The attached zip file contains " + FileList.Count.ToString() + " files:" + "\r\n";

            foreach (ChosenFile chosenFile in FileList)
            {
                string path = chosenFile.filePath;
                string fileName = chosenFile.fileName;
                string extension = Path.GetExtension(path);
                string fileNameNoExt = fileName.Substring(0, fileName.Length - extension.Length);

                MailSubject += fileNameNoExt;
                if (count < FileList.Count - 1)
                {
                    MailSubject += "_";
                }
                MailBody += fileNameNoExt;
                MailBody += "\r\n";

                count ++;
            }

            MailBody += "\r\n";
            MailBody += CustomText.Text + "\r\n";
            MailBody += "\r\n";
            MailBody += permanentMessage + "\r\n";

            //GET CREDENTIALS
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string credentialsPath = Path.Combine(appPathPrevious, @"Credentials.txt");

            if (!File.Exists(credentialsPath))
            {
                Message.Text = "Couldn't find Credentials.txt in the path: " + credentialsPath;
                WindowState = WindowState.Normal;
                return;
            }
            string[] credentialsText = System.IO.File.ReadAllLines(credentialsPath);

            string senderName = credentialsText[0].Split(':', 2)[1].Trim();
            string senderAdress = credentialsText[1].Split(':', 2)[1].Trim();
            string password = credentialsText[2].Split(':', 2)[1].Trim();

            //GET RECIPIENT ADRESS
            string recipientPath = Path.Combine(appPathPrevious, @"Recipient.txt");
            if (!File.Exists(recipientPath))
            {
                Message.Text = "Couldn't find Recipient.txt in the path: " + recipientPath;
                WindowState = WindowState.Normal;
                return;
            }

            string recipientAdress = System.IO.File.ReadAllText(recipientPath);

            NetworkCredential credentials = new NetworkCredential(senderAdress, password);

            MailMessage mail = new MailMessage()
            {
                From = new MailAddress(senderAdress, senderName),
                Subject = MailSubject,
                Body = MailBody
            };

            mail.Attachments.Add(new Attachment(zipFile));

            mail.To.Add(new MailAddress(recipientAdress));

            // Smtp client
            var client = new SmtpClient()
            {
                Port = 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Host = "smtp.gmail.com",
                EnableSsl = true,
                Credentials = credentials
            };

            try
            {
                client.Send(mail);
            }
            catch (Exception ex)
            {
                Message.Text = "Mail failed: " + ex.Message;
                WindowState = WindowState.Normal;
                return;
            }

            mail.Dispose();

            //File.Delete(zipFile);
            Message.Text = "Mail sent successfully";

            if ((bool)DeleteFilesBox.IsChecked)
            {
                foreach(ChosenFile chosenFile in FileList)
                {
                    if (File.Exists(chosenFile.filePath))
                    {
                        File.Delete(chosenFile.filePath);
                    }
                }
            }

            if(File.Exists(zipFile))
            {
                try
                {
                    File.Delete(zipFile);
                }
                catch (Exception ex)
                {
                    Message.Text = "Couldn't delete .zip File: " + ex.Message;
                    WindowState = WindowState.Normal;
                    return;
                }
            }

            System.Windows.Application.Current.Shutdown();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            int index = FileBox.SelectedIndex;
            FileList.RemoveAt(index);
            FileMessage.Text = "";
            //this.FileBox.Items.Refresh();
        }

        static string ReduceName(string path, ObservableCollection<ChosenFile> FileList)
        {
            string fileName = Path.GetFileName(path);
            string extension = Path.GetExtension(path);
            string fileNameNoExt = Path.GetFileNameWithoutExtension(path);

            string[] splitFileName = fileNameNoExt.Split('-');
            if (splitFileName.Length > 2)
            {
                fileName = splitFileName[1] + extension;
                fileNameNoExt = splitFileName[1];
            }

            bool exists = true;
            int suffix = 0;
            string ofileName = fileName;
            string ofileNameNoExt = fileNameNoExt;

            while (exists)
            {
                exists = false;

                foreach (ChosenFile file in FileList)
                {
                    if(!(0==suffix))
                    {
                        fileNameNoExt = ofileNameNoExt;
                        fileNameNoExt += "-" + suffix.ToString();
                        fileName = fileNameNoExt + extension;
                    }
                    if(file.fileName == fileName)
                    {
                        exists = true;
                        break;
                    }
                }

                suffix++;
            }

            return fileName;
        }

        static string CreateZipFolder(ObservableCollection<ChosenFile> FileList)
        {
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string zipFolder = Path.Combine(Path.GetDirectoryName(appPathPrevious), "ZipFolder");
            string zipFile = Path.Combine(Path.GetDirectoryName(appPathPrevious), "ZipFile.zip");

            if (System.IO.File.Exists(zipFile))
            {
                try
                {
                    File.Delete(zipFile);
                }
                catch (Exception ex)
                {
                    throw;
                }
            }

            if (Directory.Exists(zipFolder))
            {
                Directory.Delete(zipFolder, true);
            }

            Directory.CreateDirectory(zipFolder);

            foreach (ChosenFile chosenFile in FileList)
            {
                string path = chosenFile.filePath;

                string newPath = Path.Combine(zipFolder, chosenFile.fileName);

                try
                {
                    System.IO.File.Copy(chosenFile.filePath, newPath);
                }
                catch (Exception ex)
                {
                    throw;
                }
            }

            string[] zipFiles = Directory.GetFiles(zipFolder);

            try
            {
                string sevenZDllPath = Path.Combine(appPath, "7z.dll");
                SevenZip.SevenZipCompressor.SetLibraryPath(sevenZDllPath);
                SevenZip.SevenZipCompressor szc = new SevenZip.SevenZipCompressor();
                szc.PreserveDirectoryRoot= true;
                szc.CompressionLevel = SevenZip.CompressionLevel.Ultra;
                szc.CompressionMode = SevenZip.CompressionMode.Create;
                szc.DirectoryStructure = true;
                //szc.EncryptHeaders = true;
                szc.DefaultItemName = zipFile; //if the full path given the folders are also created
                szc.CompressFiles(zipFile, zipFiles);
            }
            catch (Exception ex)
            {
                throw;
            }

            /*
            System.IO.Compression.ZipFile.CreateFromDirectory(zipFolder, Path.Combine(Path.GetDirectoryName(appPathPrevious), "ZipFile_old.zip"), CompressionLevel.SmallestSize, false);
            */            

            long sizebytes = new FileInfo(zipFile).Length;
            float sizemb = sizebytes / 1048576;
            if(sizemb > 24.5)
            {
                return "toobig";
            }

            //DELETE THE FOLDER OF THE COPIED FILES
            try
            {
                Directory.Delete(zipFolder, true);
            }
            catch (Exception ex)
            {
                throw;
            }

            return zipFile;
        }
    }
}
