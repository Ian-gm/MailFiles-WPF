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
using static System.Net.Mime.MediaTypeNames;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using MailingWPF.Properties;

namespace MailingWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    public partial class MainWindow : Window
    {
        public ObservableCollection<ChosenFile> FileList { get; set; }

        public string[] senderData;
        public string recipientAdress;

        public MainWindow()
        {
            FileList = new ObservableCollection<ChosenFile>();
            InitializeComponent();
            this.DataContext = this;
            Rotulo.Text = "by Mutable Substance: mutable.substance@gmail.com";

            string[] args = Environment.GetCommandLineArgs(); //Si el programa fue abierto con context menu, leer el archivo seleccionado
            foreach (string s in args)
            {
                if (!Path.GetFileName(s).EndsWith(".dll"))
                {
                    string fileName = ReduceName(s, FileList);
                    FileList.Add(new ChosenFile(s, fileName));
                }
            }

            //READ ALL .TXT FILES
            string appPath = AppContext.BaseDirectory;
            string appPathPrevious = Directory.GetParent(appPath).Parent.FullName;
            string credentialsPath = Path.Combine(appPathPrevious, @"Credentials.txt");
            string recipientPath = Path.Combine(appPathPrevious, @"Recipient.txt");

            if (!File.Exists(credentialsPath))
            {
                StatusMessage.Text = "Couldn't find Credentials.txt in the path: " + credentialsPath + "\r\n" + "\r\n";
                StatusMessage.Text += "Please create the file and add the information as follows: " + "\r\n";
                StatusMessage.Text += "name: Sender's Name " + "\r\n";
                StatusMessage.Text += "from: sender's@gmail.com " + "\r\n";
                StatusMessage.Text += "password: gmail app password " + "\r\n";
                StatusMessage.Text += "\r\n";
                SendButton.IsEnabled = false;
            }
            else
            {
                try
                {
                    senderData = getCredentials(credentialsPath);
                }
                catch (Exception e)
                {
                    StatusMessage.Text = "Couldn't read the Credentials.txt, error message:" + "\r\n" + e.Message;
                    SendButton.IsEnabled = false;
                }

                StatusMessage.Text = "Credentials read successfully" + "\r\n";
            }

            if (!File.Exists(recipientPath))
            {
                StatusMessage.Text = "Couldn't find Recipient.txt in the path: " + recipientPath;
                StatusMessage.Text += "Please create the file and add the information as follows: " + "\r\n";
                StatusMessage.Text += "recipient's@gmail.com " + "\r\n";
                WindowState = WindowState.Normal;
            }
            else
            {
                try
                {
                    recipientAdress = System.IO.File.ReadAllText(recipientPath);
                }
                catch (Exception e)
                {
                    StatusMessage.Text = "Couldn't read the Recipient.txt, error message:" + "\r\n" + e.Message;
                }

                StatusMessage.Text += "Recipient's mail adress read successfully";

                Recipient.Text = recipientAdress;
            }

            string chosenFolder = Properties.Settings.Default.CopyFolder;

            if (chosenFolder != null )
            {
                ChosenFolder.Text = "Folder: " + Path.GetFileName(chosenFolder);
            }
        }

        public class ChosenFile //Este objeto guarda la información de los archivos seleccionados
        {
            public ChosenFile(string Path, string Name)
            {
                filePath = Path;
                fileName = Name;
            }
            public string filePath { get; set; }
            public string fileName { get; set; }
        }

        private void dropfiles(object sender, System.Windows.DragEventArgs e) //Esta es la función que recibe archivos por drag n drop
        {
            string[] droppedFiles = null;

            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                droppedFiles = e.Data.GetData(System.Windows.DataFormats.FileDrop, true) as string[];
            }

            if ((null == droppedFiles) || (!droppedFiles.Any())) { return; }

            foreach (string s in droppedFiles)
            {
                string fileName = ReduceName(s, FileList);

                FileList.Add(new ChosenFile(s, fileName));
            }

            SelectedFilesTotalSize();
        }

        private void SelectedFilesTotalSize()
        {
            long totalSize = 0;
            foreach (ChosenFile chosenFile in FileList)
            {
                totalSize += new FileInfo(chosenFile.filePath).Length;
            }

            double toMb = Math.Round((double)(totalSize) / 1000000, 2);
            Totalfilesize.Text = toMb.ToString() + "mb";
            if (toMb > 25)
            {
                Totalfilesize.Foreground = new SolidColorBrush(Colors.Red);
                Totalfilesize.Text += "\r\n" + "Zip file may be too large";
            }
            else
            {
                Totalfilesize.Foreground = new SolidColorBrush(Colors.Black);
            }
        }
        private string[] getCredentials(string credentialsPath)
        {
            string[] credentialsText;
            try
            {
                credentialsText = System.IO.File.ReadAllLines(credentialsPath);
            }
            catch (Exception ex)
            {
                StatusMessage.Text = "Recipient data was successfully read" + "\r\n";
                WindowState = WindowState.Normal;
                throw;
            }
            string senderName = credentialsText[0].Split(':', 2)[1].Trim();
            string senderAdress = credentialsText[1].Split(':', 2)[1].Trim();
            string password = credentialsText[2].Split(':', 2)[1].Trim();

            string[] senderData = new string[] { senderName, senderAdress, password };
            return senderData;
        }

        private void Recipient_TextChanged(object sender, TextChangedEventArgs e)
        {
           recipientAdress = Recipient.Text;
        }

        //Interface
        private void SendButton_Click(object sender, RoutedEventArgs e) //Este botón zipea los archivos y envía el mail
        {

            foreach (ChosenFile chosenFile in FileList)
            {
                int sameCount = 0;
                string fileName = chosenFile.fileName;
                foreach (ChosenFile otherFile in FileList)
                {
                    if(fileName == otherFile.fileName)
                    {
                        sameCount++;
                    }
                }
                if (sameCount > 1)
                {
                    StatusMessage.Text = "Some selected files are named the same, please rename them in order to continue";
                    return;
                }
            }




            WindowState = WindowState.Minimized;

            string zipFile = "";

            //ZIP
            try
            {
                zipFile = CreateZipFolder(FileList);
            }
            catch(Exception exep)
            {
                StatusMessage.Text = $"Zip Failed: {exep}";
                WindowState = WindowState.Normal;
                return;
            }

            if (zipFile == "toobig")
            {
                StatusMessage.Text = "The Zip File is larger than 25mb";
                WindowState = WindowState.Normal;
                return;
            }

            //MAIL SUBJECT AND BODY
            string MailBody = "";
            string MailSubject = "";
            string permanentMessage = "If any file is denoted with 'HR' in the filename, it is intended for High-Resolution print.";
            int count = 0;

            MailBody += CustomText.Text + "\r\n";
            MailBody += "\r\n";

            MailBody += "The attached zip file contains " + FileList.Count.ToString() + " files:" + "\r\n";

            foreach (ChosenFile chosenFile in FileList)
            {
                string path = chosenFile.filePath;
                string fileName = chosenFile.fileName;

                MailSubject += fileName;
                if (count < FileList.Count - 1)
                {
                    MailSubject += "_";
                }
                MailBody += fileName;
                MailBody += "\r\n";

                count++;
            }

            MailBody += "\r\n";
            MailBody += permanentMessage + "\r\n";


            //CREATE MAIL
            string senderName = senderData[0];
            string senderAdress = senderData[1];
            string password = senderData[2];

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
                StatusMessage.Text = "Mail failed: " + ex.Message;
                WindowState = WindowState.Normal;
                return;
            }

            mail.Dispose();

            StatusMessage.Text = "Mail sent successfully";

            //COPIAR ARCHIVOS
            if ((bool)CopyFilesBox.IsChecked)
            {
                string Folder = Properties.Settings.Default.CopyFolder;

                if(Folder == null)
                {
                    string appPath = AppContext.BaseDirectory;
                    string rootPath = Directory.GetParent(appPath).Root.FullName;
                    Folder = Path.Combine(rootPath, "SelectedFiles");
                    Properties.Settings.Default.CopyFolder = Folder;
                    Properties.Settings.Default.Save();
                }

                if(!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }

                foreach (ChosenFile chosenFile in FileList)
                {
                    if (File.Exists(chosenFile.filePath))
                    {
                        string fileName = Path.GetFileName(chosenFile.filePath);
                        string finalPath = Path.Combine(Folder, fileName);
                        try
                        {
                            File.Copy(chosenFile.filePath, finalPath);
                        }
                        catch (IOException ex)
                        {
                            StatusMessage.Text = ex.Message;
                        }
                    }
                }

                WindowState = WindowState.Normal;
                System.Windows.MessageBox.Show("Files were copied to: " + Folder.ToString());
            }

            //BORRAR ARCHIVOS
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
                    StatusMessage.Text = "Couldn't delete .zip File: " + ex.Message;
                    WindowState = WindowState.Normal;
                }
            }

            System.Windows.Application.Current.Shutdown();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            int index = FileBox.SelectedIndex;
            try
            {
                FileList.RemoveAt(index);
            }
            catch(Exception ex)
            {
                StatusMessage.Text = "Failed to remove file from list, are you sure you clicked on a file before clicking on this button?";
            }

            SelectedFilesTotalSize();
        }

        private void ChooseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                string chosenFolder = fbd.SelectedPath;
                Properties.Settings.Default.CopyFolder = chosenFolder;
                Properties.Settings.Default.Save();

                ChosenFolder.Text = "Folder: " + Path.GetFileName(chosenFolder);
                StatusMessage.Text = "Chosen folder to copy the selected files: " + chosenFolder;
            }
        }

        //RECORTA EL NOMBRE DE LOS ARCHIVOS
        static string ReduceName(string path, ObservableCollection<ChosenFile> FileList) 
        {
            string prevfileName = null;
            string fileName;
            string fileNameNoExt = Path.GetFileNameWithoutExtension(path);

            string[] splitFileName = fileNameNoExt.Split('-');
            if (splitFileName.Length > 2)
            {
                prevfileName = splitFileName[0];
                fileName = splitFileName[1];
            }
            else if (fileNameNoExt.Length > 6)
            {
                fileName = fileNameNoExt.Substring(0,6);
            }
            else
            {
                fileName = fileNameNoExt;
            }

            bool exists = true;
            int suffix = -1;
            string ofileName = fileName;

            while (exists)
            {
                exists = false;

                foreach (ChosenFile file in FileList)
                {
                    if (suffix == -1) 
                    {
                        fileName = ofileName;
                    }
                    else if (suffix == 0) 
                    {
                        if (prevfileName != null)
                        {
                            fileName = ofileName;
                            fileName = prevfileName + "-" + fileName;
                        }
                    }
                    else
                    {
                        fileName = ofileName;
                        if (prevfileName != null)
                        {
                            fileName = prevfileName + "-" + fileName;
                        }
                        fileName += " (" + suffix.ToString() + ")";
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
                string newPath = Path.Combine(zipFolder, chosenFile.fileName + Path.GetExtension(chosenFile.filePath));

                try
                {
                    System.IO.File.Copy(path, newPath);
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