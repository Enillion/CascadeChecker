using System;
using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs; //Uber solution for browsing folders
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace WPFCascadeCheckerUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();            
        }

        // CascadeChecker Tab Elements
        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fileBrowser = new CommonOpenFileDialog
            {
                Multiselect = false,
                IsFolderPicker = true
            };

            if (fileBrowser.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string browsedPath = fileBrowser.FileName;
                PathDisplay.Text = browsedPath;
            }
        }

        private void StackDropFolder_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                PathDisplay.Text = path[0];
            }
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            
            if (PathDisplay.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                //-------------------------------> disable wpf controls
                RunButton.IsEnabled = false;
                LogButton.IsEnabled = false;
                LogLocation.IsEnabled = false;
                StackDropFolder.IsEnabled = false;
                PathDisplay.IsEnabled = false;
                BrowseFolder.IsEnabled = false;
                UICOverride.IsEnabled = false;
                CheckXmlLang.IsEnabled = false;
                AddXmlLang.IsEnabled = false;

                progressBar1.Value = 0;
                progressBar1.Maximum = (Directory.GetFiles(PathDisplay.Text, "*", SearchOption.AllDirectories)).Length -1;
                
                string logFilePath = PathDisplay.Text.Substring(0, PathDisplay.Text.LastIndexOf(@"\")) + @"\Checker_Log_" + 
                    PathDisplay.Text.Substring(PathDisplay.Text.LastIndexOf(@"\") + 1, (PathDisplay.Text.Length - PathDisplay.Text.LastIndexOf(@"\") - 1)) + "_" + DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss") + ".txt";

                LogLocation.Text = logFilePath;
                string uicOverride = "standard";
                if (UICOverride.IsChecked == true)
                {
                    uicOverride = "simplified";
                }
                string[] arguments = {PathDisplay.Text, logFilePath, uicOverride};                

                BackgroundWorker bgw = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw.WorkerReportsProgress = true;
                bgw.DoWork += Bgw_DoWork;
                bgw.ProgressChanged += Bgw_ProgressChanged;
                bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted;
                bgw.RunWorkerAsync(arguments);                
            }
        }//-----RunButton_Click section start

        private void Bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Check Finished\nLog File created.");
            //-------------------------------> enable wpf controls
            RunButton.IsEnabled = true;
            LogButton.IsEnabled = true;
            LogLocation.IsEnabled = true;
            StackDropFolder.IsEnabled = true;
            PathDisplay.IsEnabled = true;
            BrowseFolder.IsEnabled = true;
            UICOverride.IsEnabled = true;
            CheckXmlLang.IsEnabled = true;
            AddXmlLang.IsEnabled = true;
        }

        private void Bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void Bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] arguments = (string[])e.Argument;
            string logFilePath = arguments[1];
            string pathDisplay = arguments[0];
            string uicOverride = arguments[2];
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                FileInfo fi = new FileInfo(logFilePath);
                //Create text file using "fi" path
                using (StreamWriter sw = fi.CreateText())
                {
                    sw.WriteLine("Issues found (if any):");
                    string[] files = Directory.GetFiles(pathDisplay, "*", SearchOption.AllDirectories);

                    foreach (string file in files)
                    {
                        XmlTools.CheckXml(file, sw, uicOverride);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                    }                    
                }                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }//-----RunButton_Click section end

        private void LogButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(LogLocation.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("Log file for this check does not exist");
            }
        }

        private void AddXmlLang_Click(object sender, RoutedEventArgs e)//-----AddXmlLang_Click section start
        {
            if (PathDisplay.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                //-------------------------------> disable wpf controls
                RunButton.IsEnabled = false;
                LogButton.IsEnabled = false;
                LogLocation.IsEnabled = false;
                StackDropFolder.IsEnabled = false;
                PathDisplay.IsEnabled = false;
                BrowseFolder.IsEnabled = false;
                UICOverride.IsEnabled = false;
                CheckXmlLang.IsEnabled = false;
                AddXmlLang.IsEnabled = false;

                progressBar1.Value = 0;
                progressBar1.Maximum = (Directory.GetFiles(PathDisplay.Text, "*", SearchOption.AllDirectories)).Length - 1;
                                
                string[] arguments = { PathDisplay.Text };

                BackgroundWorker bgw = new BackgroundWorker();
                bgw.WorkerReportsProgress = true;
                bgw.DoWork += Bgw_DoWork2;
                bgw.ProgressChanged += Bgw_ProgressChanged2;
                bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted2;
                bgw.RunWorkerAsync(arguments);
            }
        }

        private void Bgw_RunWorkerCompleted2(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Done.");
            //-------------------------------> enable wpf controls
            RunButton.IsEnabled = true;
            LogButton.IsEnabled = true;
            LogLocation.IsEnabled = true;
            StackDropFolder.IsEnabled = true;
            PathDisplay.IsEnabled = true;
            BrowseFolder.IsEnabled = true;
            UICOverride.IsEnabled = true;
            CheckXmlLang.IsEnabled = true;
            AddXmlLang.IsEnabled = true;
        }

        private void Bgw_ProgressChanged2(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void Bgw_DoWork2(object sender, DoWorkEventArgs e)
        {
            string[] arguments = (string[])e.Argument;            
            string pathDisplay = arguments[0];
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                //string[] files = Directory.GetFiles(pathDisplay, "*", SearchOption.AllDirectories);
                string[] folders = Directory.GetDirectories(pathDisplay);


                foreach (string folder in folders)
                {
                    string[] files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
                    XmlAttribute xmlLang = null;

                    foreach (string file in files)
                    {                      
                        if (Path.GetExtension(file) == ".xml" || Path.GetExtension(file) == ".ditamap" || Path.GetExtension(file) == ".dita")
                        {
                            try
                            {
                                XmlDocument xml = new XmlDocument();

                                using (XmlReader xr = new XmlTextReader(file) { Namespaces = false })
                                {
                                    xml = new PositionXmlDocument();
                                    xml.Load(xr);
                                }

                                var root = xml.DocumentElement;
                                string missingXmlLangPath = file.Substring(0, file.LastIndexOf('\\'));
                                if (root.HasAttribute("xml:lang") == true)
                                {
                                    xmlLang = root.GetAttributeNode("xml:lang");                                    
                                    if (xmlLang.Value == "nb_NO")
                                    {
                                        xmlLang.Value = "no_NO";
                                    }
                                    else if (xmlLang.Value == "es-CO")
                                    {
                                        xmlLang.Value = "es_CO";
                                    }
                                    else if (xmlLang.Value == "sr_RS_Latn")
                                    {
                                        xmlLang.Value = "sr_RS";
                                    }
                                    continue;
                                }
                            }
                            catch (Exception exc)
                            {
                                MessageBox.Show(exc.Message);
                            }
                        }                     
                    }
                    foreach (string file in files)
                    {
                        if (xmlLang != null)
                        {
                            if (Path.GetExtension(file) == ".xml" || Path.GetExtension(file) == ".ditamap")
                            {
                                try
                                {
                                    XmlDocument xml = new XmlDocument();

                                    using (XmlReader xr = new XmlTextReader(file) { Namespaces = false })
                                    {
                                        xml = new PositionXmlDocument();
                                        xml.Load(xr);
                                    }

                                    if (xml.DocumentElement.HasAttribute("xml:lang") == false || xml.DocumentElement.GetAttributeNode("xml:lang").Value == "nb_NO" || xml.DocumentElement.GetAttributeNode("xml:lang").Value == "es-CO" || xml.DocumentElement.GetAttributeNode("xml:lang").Value == "sr_RS_Latn")
                                    {
                                        if (xml.DocumentElement.HasAttribute("xml:lang") == true)
                                        {
                                            xml.DocumentElement.RemoveAttribute("xml:lang");
                                        }
                                        XmlAttribute newXmlLang = xml.CreateAttribute("xml:lang");                                        
                                        newXmlLang.Value = xmlLang.Value;
                                        xml.DocumentElement.SetAttributeNode(newXmlLang);                                       
                                        string fileContent;//-----------------------------Output xml must be UTF8 without bom with Unix LF as opposed to NET defaults so below workaround is needed
                                        using (StringWriter sw = new StringWriter())
                                        {
                                            using (XmlTextWriter tx = new XmlTextWriter(sw))
                                            {
                                                xml.WriteTo(tx);
                                            }
                                            fileContent = sw.ToString();
                                        }
                                            
                                        StringBuilder sb = new StringBuilder();
                                        foreach (char character in fileContent)
                                        {
                                            sb.Append(character);
                                        }
                                        //sb.Replace(@"\r\n", @"\n");//-----------convert Windows newline with Unix newline
                                        sb.Replace(Environment.NewLine, ('\u000A').ToString());//-----------convert Windows newline with Unix newline
                                        File.WriteAllText(file, sb.ToString());                                        
                                    }                                                
                                }
                                catch (Exception exc)
                                {
                                    MessageBox.Show(exc.Message + "\n" + exc.StackTrace);
                                }
                            }
                        }
                        (sender as BackgroundWorker).ReportProgress(progress++);
                    }
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }

        }//-----AddXmlLang_Click section end

        private void CheckXmlLang_Click(object sender, RoutedEventArgs e)
        {
            if (PathDisplay.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                //-------------------------------> disable wpf controls
                RunButton.IsEnabled = false;
                LogButton.IsEnabled = false;
                LogLocation.IsEnabled = false;
                StackDropFolder.IsEnabled = false;
                PathDisplay.IsEnabled = false;
                BrowseFolder.IsEnabled = false;
                UICOverride.IsEnabled = false;
                CheckXmlLang.IsEnabled = false;
                AddXmlLang.IsEnabled = false;

                progressBar1.Value = 0;
                progressBar1.Maximum = (Directory.GetFiles(PathDisplay.Text, "*", SearchOption.AllDirectories)).Length - 1;

                string logFilePath = PathDisplay.Text.Substring(0, PathDisplay.Text.LastIndexOf(@"\")) + @"\Checker_XMLLANG_Log_" +
                    PathDisplay.Text.Substring(PathDisplay.Text.LastIndexOf(@"\") + 1, (PathDisplay.Text.Length - PathDisplay.Text.LastIndexOf(@"\") - 1)) + "_" + DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss") + ".txt";

                LogLocation.Text = logFilePath;

                string[] arguments = { PathDisplay.Text, logFilePath };

                BackgroundWorker bgw = new BackgroundWorker();
                bgw.WorkerReportsProgress = true;
                bgw.DoWork += Bgw_DoWork1;
                bgw.ProgressChanged += Bgw_ProgressChanged1;
                bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted1;
                bgw.RunWorkerAsync(arguments);                
            }
        }//-----CheckXmlLang_Click section start

        private void Bgw_RunWorkerCompleted1(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Check complete.\nLog created.");
            //-------------------------------> enable wpf controls
            RunButton.IsEnabled = true;
            LogButton.IsEnabled = true;
            LogLocation.IsEnabled = true;
            StackDropFolder.IsEnabled = true;
            PathDisplay.IsEnabled = true;
            BrowseFolder.IsEnabled = true;
            UICOverride.IsEnabled = true;
            CheckXmlLang.IsEnabled = true;
            AddXmlLang.IsEnabled = true;
        }

        private void Bgw_ProgressChanged1(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void Bgw_DoWork1(object sender, DoWorkEventArgs e)
        {
            string[] arguments = (string[])e.Argument;
            string logFilePath = arguments[1];
            string pathDisplay = arguments[0];            
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                FileInfo fi = new FileInfo(logFilePath);
                //Create text file using "fi" path
                using (StreamWriter sw = fi.CreateText())
                {
                    sw.WriteLine("Language folders with missing xml:lang attributes (if any):");
                    string[] files = Directory.GetFiles(pathDisplay, "*", SearchOption.AllDirectories);
                    string filePath = "";
                    foreach (string file in files)
                    {
                        if (Path.GetExtension(file) == ".xml" || Path.GetExtension(file) == ".ditamap" || Path.GetExtension(file) == ".dita")
                        {                            
                            try
                            {
                                XmlDocument xml = new XmlDocument();

                                using (XmlReader xr = new XmlTextReader(file) { Namespaces = false })
                                {
                                    xml = new PositionXmlDocument();
                                    xml.Load(xr);
                                }
                                
                                var root = xml.DocumentElement;
                                string missingXmlLangPath = file.Substring(0, file.LastIndexOf('\\'));
                                if (root.HasAttribute("xml:lang") == false)
                                {
                                    if (filePath != missingXmlLangPath)
                                    {
                                        sw.WriteLine("\n\n{0} - Some files are missing xml:lang attribute", missingXmlLangPath);
                                        filePath = missingXmlLangPath;
                                    }                                    
                                }
                            }
                            catch (Exception exc)
                            {
                                if (filePath != file)
                                {
                                    sw.WriteLine("\n\n\nFile: {0}", file);
                                    filePath = file;
                                }
                                sw.WriteLine("    {0}\n", exc.Message);
                            }
                        }                        
                        (sender as BackgroundWorker).ReportProgress(progress++);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }//-----CheckXmlLang_Click section end

        //XLIFF Batcher Tab Elements
        private void BrowseFolder2_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fileBrowser = new CommonOpenFileDialog
            {
                Multiselect = false,
                IsFolderPicker = true
            };

            if (fileBrowser.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string browsedPath = fileBrowser.FileName;
                PathDisplay2.Text = browsedPath;
            }
        }

        private void StackDropFolder2_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                PathDisplay2.Text = path[0];
            }
        }

        private void CreateFoldersButton_Click(object sender, RoutedEventArgs e)
        {
            if (PathDisplay2.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                try
                {
                    BatcherStack.IsEnabled = false;

                    BatchProgress.Value = 0;
                    //BatchProgress.Maximum = (Directory.GetDirectories(PathDisplay2.Text)).Length + 1;
                    BatchProgress.Maximum = (Directory.GetFiles(PathDisplay2.Text)).Length + 1;

                    BackgroundWorker bgwBatch = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                    bgwBatch.WorkerReportsProgress = true;
                    bgwBatch.DoWork += BgwBatch_DoWork;
                    bgwBatch.ProgressChanged += BgwBatch_ProgressChanged;
                    bgwBatch.RunWorkerCompleted += BgwBatch_RunWorkerCompleted;
                    bgwBatch.RunWorkerAsync(PathDisplay2.Text);                                       
                }
                catch (Exception sKex)
                {
                    MessageBox.Show(sKex.Message + "\n" + sKex.StackTrace);
                }                
            }
        }

        private void BgwBatch_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            BatcherStack.IsEnabled = true;
        }

        private void BgwBatch_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BatchProgress.Value = e.ProgressPercentage;
        }

        private void BgwBatch_DoWork(object sender, DoWorkEventArgs e)
        {
            string filePath = (string)e.Argument;
            int progress = 0;

            try
            {
                if (Directory.Exists(filePath) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                //Create array of filenames with path from directory and subfolders
                string[] files = Directory.GetFiles(filePath, "*", SearchOption.AllDirectories);

                foreach (var file in files)
                {
                    if (Path.GetExtension(file) == ".sdlxliff") //only sdlxliff files exported frpm pasolo are processed
                    {
                        string languageCode = file.Substring((file.LastIndexOf("_") + 1), ((file.Length - 10) - file.LastIndexOf("_"))).ToLower();
                        languageCode = SwissKnifeMethods.CheckLanguageCode(languageCode);
                        string foldername = filePath + @"\" + languageCode;
                        if (Directory.Exists(foldername))
                        {
                            MessageBoxButton buttons = MessageBoxButton.YesNo;

                            MessageBoxResult result = MessageBox.Show("Cannot create folder as it already exists\nDo you wish to continue anyway?", languageCode, buttons);
                            if (result == MessageBoxResult.No)
                            {
                                return;
                            }
                            else
                            {
                                continue;
                            }
                        }
                        try
                        {
                            //Folder creation and moving file from default location to the new folder
                            DirectoryInfo di = Directory.CreateDirectory(foldername);
                            File.Move(file, foldername + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\"))));
                        }
                        catch (Exception dirFile)
                        {
                            MessageBox.Show(dirFile.Message);
                            continue;
                        }
                    }
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }
            }
            catch (IOException ioe)
            {
                MessageBox.Show(ioe.Message);
                return;
            }
            catch (UnauthorizedAccessException uae)
            {
                MessageBox.Show(uae.Message);
                return;
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message);
                return;
            }

            try
            {
                XmlTools.CreateConfigFile(filePath);
                (sender as BackgroundWorker).ReportProgress(progress++);
            }
            catch (Exception configException)
            {
                MessageBox.Show("Error during config generation: " + configException.Message);
                return;
            }

            try
            {
                //ZipFile.CreateFromDirectory returns exception when trying to create .zip file in the same directory which contents being zipped, therefore new directory need to be created.
                string outputFolder = filePath.Substring(0, filePath.LastIndexOf(@"\")) + @"\TMS_Batch_Folder";
                DirectoryInfo di = Directory.CreateDirectory(outputFolder);
                ZipFile.CreateFromDirectory(filePath, (outputFolder + @"\TMS_Batch.zip"));
                (sender as BackgroundWorker).ReportProgress(progress++);
            }
            catch (Exception zipex)
            {
                MessageBox.Show(zipex.Message);
            }
            MessageBox.Show("Done");
        }

        //Native 2 ASCII TAB Elements
        private void BrowseFolder3_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fileBrowser = new CommonOpenFileDialog
            {
                Multiselect = false,
                IsFolderPicker = true
            };

            if (fileBrowser.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string browsedPath = fileBrowser.FileName;
                PathDisplay3.Text = browsedPath;
            }
        }

        private void StackDropFolder3_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                PathDisplay3.Text = path[0];
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            if (PathDisplay3.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                Convert.IsEnabled = false;
                ChangeEncoding.IsEnabled = false;
                StackDropFolder3.IsEnabled = false;
                PathDisplay3.IsEnabled = false;
                BrowseFolder3.IsEnabled = false;
                ApoCheck.IsEnabled = false;

                progressBar2.Value = 0;
                progressBar2.Maximum = (Directory.GetFiles(PathDisplay3.Text, "*", SearchOption.AllDirectories)).Length - 1;

                string apoCheck = "unchecked";
                if (ApoCheck.IsChecked == true)
                {
                    apoCheck = "checked";
                }
                string[] convertArguments = { PathDisplay3.Text, apoCheck };

                BackgroundWorker bgw2 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw2.WorkerReportsProgress = true;
                bgw2.DoWork += Bgw2_DoWork;
                bgw2.ProgressChanged += Bgw2_ProgressChanged;
                bgw2.RunWorkerCompleted += Bgw2_RunWorkerCompleted;
                bgw2.RunWorkerAsync(convertArguments);
            }
        }

        private void Bgw2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Conversion Complete");
            //-------------------------------> enable wpf controls
            Convert.IsEnabled = true;
            ChangeEncoding.IsEnabled = true;
            StackDropFolder3.IsEnabled = true;
            PathDisplay3.IsEnabled = true;
            BrowseFolder3.IsEnabled = true;
            ApoCheck.IsEnabled = true;
        }

        private void Bgw2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar2.Value = e.ProgressPercentage;
        }

        private void Bgw2_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] convertArguments = (string[])e.Argument;
            string directory = convertArguments[0];
            string apoCheck = convertArguments[1];

            int progress = 0;

            try
            {
                if (Directory.Exists(directory) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }
                
                string[] files = Directory.GetFiles(directory, "*", SearchOption.AllDirectories);
                                                                                 
                foreach (string file in files)
                {
                    string fileName = file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\")));
                    using (StreamReader reader = new StreamReader(file))
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(reader.ReadToEnd());
                        string fileContent = sb.ToString();
                        sb.Clear();
                        foreach (char character in fileContent)
                        {
                            if (character > 127) //if character position in encodepage exceeds 127 ASCII character set
                            {
                                string escapedChar = string.Format(@"\u{0:x4}", (int)character);
                                sb.Append(escapedChar);
                            }
                            else
                            {
                                if (character == '\'' & apoCheck == "checked")
                                {
                                    string escapedApostrophe = @"\u0027";
                                    sb.Append(escapedApostrophe);                                    
                                }
                                else
                                {
                                    sb.Append(character);
                                }
                            }
                        }
                        string filePath = file.Substring(directory.Length + 1, (file.Length - (directory.Length + fileName.Length))); // path to the file without path to the directory                        
                        string outputFolder = "";//path to the new directory with converted files
                        if (filePath == "" || filePath[0] == '\\') // in case file is directly placed in folder chosen for conversion
                        {
                            outputFolder = directory.Substring(0, directory.LastIndexOf(@"\")) + @"\Converted Files" + filePath;
                        }
                        else
                        {
                            outputFolder = directory.Substring(0, directory.LastIndexOf(@"\")) + @"\Converted Files\" + filePath;
                        }
                        DirectoryInfo di = Directory.CreateDirectory(outputFolder);
                        File.WriteAllText((outputFolder + fileName), sb.ToString());
                    }                    
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ChangeEncoding_Click(object sender, RoutedEventArgs e)
        {            
            if (PathDisplay3.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                Convert.IsEnabled = false;
                ChangeEncoding.IsEnabled = false;
                EncodingSelection.IsEnabled = false;
                StackDropFolder3.IsEnabled = false;
                PathDisplay3.IsEnabled = false;
                BrowseFolder3.IsEnabled = false;

                progressBar2.Value = 0;
                progressBar2.Maximum = (Directory.GetFiles(PathDisplay3.Text, "*", SearchOption.AllDirectories)).Length - 1;

                string[] parameters = { PathDisplay3.Text, EncodingSelection.SelectedValue.ToString() }; //Parameters will be passed to RunWorkerAsync

                BackgroundWorker bgw3 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw3.WorkerReportsProgress = true;
                bgw3.DoWork += Bgw3_DoWork;
                bgw3.ProgressChanged += Bgw3_ProgressChanged; ;
                bgw3.RunWorkerCompleted += Bgw3_RunWorkerCompleted; ;
                bgw3.RunWorkerAsync(parameters);                
            }
        }

        private void Bgw3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Conversion Complete");
            //-------------------------------> enable wpf controls
            Convert.IsEnabled = true;
            ChangeEncoding.IsEnabled = true;
            EncodingSelection.IsEnabled = true;
            StackDropFolder3.IsEnabled = true;
            PathDisplay3.IsEnabled = true;
            BrowseFolder3.IsEnabled = true;
        }

        private void Bgw3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar2.Value = e.ProgressPercentage;
        }

        private void Bgw3_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] parameters = (string[])e.Argument;
            string directory = parameters[0];
            string selection = parameters[1];
            int progress = 0;         
            
            try
            {
                if (Directory.Exists(directory) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                string[] files = Directory.GetFiles(directory, "*", SearchOption.AllDirectories);

                foreach (string file in files)
                {
                    string fileName = file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\")));
                    using (StreamReader reader = new StreamReader(file))
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(reader.ReadToEnd());
                        string fileContent = sb.ToString();
                        sb.Clear();
                        foreach (char character in fileContent)
                        {
                            sb.Append(character);
                        }
                        string filePath = file.Substring(directory.Length + 1, (file.Length - (directory.Length + fileName.Length))); // path to the file without path to the directory                        
                        string outputFolder = "";//path to the new directory with converted files
                        if (filePath == "" || filePath[0] == '\\') // in case file is directly placed in folder chosen for conversion
                        {
                            outputFolder = directory.Substring(0, directory.LastIndexOf(@"\")) + @"\With Changed Encoding" + filePath;
                        }
                        else
                        {
                            outputFolder = directory.Substring(0, directory.LastIndexOf(@"\")) + @"\With Changed Encoding\" + filePath;
                        }
                        DirectoryInfo di = Directory.CreateDirectory(outputFolder);

                        if (selection == "System.Windows.Controls.ComboBoxItem: UTF-8")
                        {
                            File.WriteAllText((outputFolder + fileName), sb.ToString());
                        }
                        else if(selection == "System.Windows.Controls.ComboBoxItem: UTF-8-BOM")
                        {                            
                            Encoding utfBom = new UTF8Encoding(true, true);//Encoding object will be used as parameter for File.WriteAllText method                           
                            File.WriteAllText((outputFolder + fileName), sb.ToString(), utfBom);
                        }
                        else
                        {
                            MessageBox.Show("Encoding selection error");
                        }                      
                    }
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        //Native 2 OTHER TOOLS TAB Elements
        private void BrowseFolder4_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog fileBrowser = new CommonOpenFileDialog
            {
                Multiselect = false,
                IsFolderPicker = true
            };

            if (fileBrowser.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string browsedPath = fileBrowser.FileName;
                PathDisplay4.Text = browsedPath;
            }
        }

        private void StackDropFolder4_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                PathDisplay4.Text = path[0];
            }
        }

        private void ZipContents_Click(object sender, RoutedEventArgs e)//-----ZipContents_Click section start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;
                
                progressBar4.Value = 0;                
                progressBar4.Maximum = (Directory.GetDirectories(PathDisplay4.Text)).Length -1;
                
                BackgroundWorker bgw4 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw4.WorkerReportsProgress = true;
                bgw4.DoWork += Bgw4_DoWork;
                bgw4.ProgressChanged += Bgw4_ProgressChanged;
                bgw4.RunWorkerCompleted += Bgw4_RunWorkerCompleted;
                bgw4.RunWorkerAsync(PathDisplay4.Text);
            }
        }

        private void Bgw4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {            
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
        }

        private void Bgw4_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw4_DoWork(object sender, DoWorkEventArgs e)
        {            
            string pathDisplay = (string)e.Argument;            
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {                    
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }
                
                string foldername = pathDisplay + @"\" + "Zipped_Subfolders_Contents";                
                if (Directory.Exists(foldername))
                {
                    MessageBoxButton buttons = MessageBoxButton.YesNo;

                    MessageBoxResult result = MessageBox.Show("Cannot create folder as it already exists\nDo you wish to overwrite it?", foldername, buttons);
                    if (result == MessageBoxResult.No)
                    {
                        MessageBox.Show("Process Aborted.");
                        return;
                    }
                    else
                    {
                        Directory.Delete(foldername, true);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                    }
                }

                string[] folders = Directory.GetDirectories(pathDisplay);
                DirectoryInfo di = Directory.CreateDirectory(foldername);
                foreach (string folder in folders)
                {
                    ZipFile.CreateFromDirectory(folder, (foldername + @"\" + folder.Substring(folder.LastIndexOf(@"\"), (folder.Length - folder.LastIndexOf(@"\"))) + @".zip"));
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }
                MessageBox.Show("Done.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ChpFix_Click(object sender, RoutedEventArgs e)//-----CHP Fix section start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;

                progressBar4.Value = 0;
                progressBar4.Maximum = (Directory.GetDirectories(PathDisplay4.Text)).Length - 1;

                BackgroundWorker bgw5 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw5.WorkerReportsProgress = true;
                bgw5.DoWork += Bgw5_DoWork;
                bgw5.ProgressChanged += Bgw5_ProgressChanged;
                bgw5.RunWorkerCompleted += Bgw5_RunWorkerCompleted;
                bgw5.RunWorkerAsync(PathDisplay4.Text);
            }
        }

        private void Bgw5_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw5_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw5_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                string[] folders = Directory.GetDirectories(pathDisplay);
                foreach (string folder in folders)
                {
                    string[] files = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
                    foreach (string file in files)
                    {
                        if (Path.GetExtension(file) == ".txt")
                        {
                            string fileContentWithQuotes = File.ReadAllText(file);
                            string fileContent = DeleteQuotes(fileContentWithQuotes);
                            StringBuilder sb = new StringBuilder(Regex.Replace(fileContent, @"(\d{1,2})-(\d{1,2})-(\d{4})", @"$3-$2-$1"));
                                                        
                            while (sb.ToString().Contains(" \","))
                            {
                                sb.Replace(" \",", "\",");                                
                            }

                            sb.Replace(Environment.NewLine, ('\u000A').ToString());//-----------convert Windows newline with Unix newline
                            string test = sb.ToString();
                            File.WriteAllText(file, sb.ToString());
                        }                        
                    }

                    string oldFolder = ReplaceLanguageCode(folder, "nb_NO", "no_no");
                    
                    string tempFolderName = oldFolder.ToLowerInvariant() + "_Temp";
                    Directory.Move(folder, tempFolderName);
                    Directory.Move(tempFolderName, oldFolder.ToLowerInvariant());
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private string ReplaceLanguageCode(string folder, string sourceLanguageCodeToReplace, string targetLanguageCode)
        {
            string oldFolder;
            int startIndex = folder.LastIndexOf('\\') + 1;
            if (folder.Substring(startIndex, 5) == sourceLanguageCodeToReplace)
            {
                oldFolder = folder.Remove(startIndex, 5) + targetLanguageCode;
            }
            else
            {
                oldFolder = folder;
            }

            return oldFolder;
        }

        private void Rename_Click(object sender, RoutedEventArgs e)//-----Rename Files section start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;

                progressBar4.Value = 0;
                progressBar4.Maximum = (Directory.GetDirectories(PathDisplay4.Text)).Length - 1;

                BackgroundWorker bgw6 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw6.WorkerReportsProgress = true;
                bgw6.DoWork += Bgw6_DoWork;
                bgw6.ProgressChanged += Bgw6_ProgressChanged;
                bgw6.RunWorkerCompleted += Bgw6_RunWorkerCompleted;
                bgw6.RunWorkerAsync(PathDisplay4.Text);
            }
        }

        private void Bgw6_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw6_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw6_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;
            int progress = 0;

            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                string[] folders = Directory.GetDirectories(pathDisplay);
                foreach (string folderPath in folders)
                {
                    string exportSummaryPath = folderPath + @"\translation_export_summary.xml";
                    XmlDocument exportSummary = new XmlDocument();
                    exportSummary.Load(exportSummaryPath);

                    XmlNode parent = exportSummary.SelectSingleNode("translationObjectFile/translationObjectProperties");

                    foreach (XmlNode child in parent)
                    {
                        var uniqueID = child.Attributes["fileUniqueID"];
                        var nodePath = child.Attributes["nodePath"];
                        if (uniqueID != null && nodePath != null)
                        {
                            string fileUniqueID = uniqueID.Value;
                            string fileOriginalName = (string)nodePath.Value.Substring(nodePath.Value.LastIndexOf(@"/") + 1);

                            Console.WriteLine("{0}    Original file name:   {1}", fileUniqueID, fileOriginalName);

                            RenameFile(fileUniqueID, fileOriginalName, folderPath);                            
                        }
                    }
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void RenameFile(string fileUniqueID, string fileOriginalName, string folderPath)
        {            
            if (fileOriginalName != "")
            {
                if (File.Exists(folderPath + @"/" + fileUniqueID + @".xml"))
                {
                    File.Move(folderPath + @"/" + fileUniqueID + @".xml", folderPath + @"/" + fileOriginalName);
                }
                else if (File.Exists(folderPath + @"/" + fileUniqueID + @".ditamap"))
                {
                    File.Move(folderPath + @"/" + fileUniqueID + @".ditamap", folderPath + @"/" + fileOriginalName);
                }
                else if (File.Exists(folderPath + @"/" + fileUniqueID + @".dita"))
                {
                    File.Move(folderPath + @"/" + fileUniqueID + @".dita", folderPath + @"/" + fileOriginalName);
                }
            }
        }

        private void ChpPrep_Click(object sender, RoutedEventArgs e)//-----CHP Prep section Start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;

                progressBar4.Value = 0;
                progressBar4.Maximum = 7;

                BackgroundWorker bgw7 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw7.WorkerReportsProgress = true;
                bgw7.DoWork += Bgw7_DoWork;
                bgw7.ProgressChanged += Bgw7_ProgressChanged; ;
                bgw7.RunWorkerCompleted += Bgw7_RunWorkerCompleted;
                bgw7.RunWorkerAsync(PathDisplay4.Text);
            }
        }

        private void Bgw7_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw7_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw7_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder containing zipped source file.");
                    return;
                }

                string[] files = Directory.GetFiles(pathDisplay);
                foreach (string file in files)
                {
                    if (Path.GetExtension(file) == ".zip")
                    {
                        string foldername = pathDisplay + @"\" + "XTM_Source";
                        if (Directory.Exists(foldername))
                        {
                            MessageBoxButton buttons = MessageBoxButton.YesNo;

                            MessageBoxResult result = MessageBox.Show("Cannot create folder \"XTM_Source\" as it already exists\nDo you wish to overwrite it?", foldername, buttons);
                            if (result == MessageBoxResult.No)
                            {
                                MessageBox.Show("Process Aborted.");
                                return;
                            }
                            else
                            {
                                Directory.Delete(foldername, true);                                
                            }
                        }

                        DirectoryInfo di = Directory.CreateDirectory(foldername);
                        string htmlFolder = foldername + @"\" + "HTML";
                        string htmlSourceFolder = htmlFolder + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));
                        string txtFolder = foldername + @"\" + "TXT";
                        string txtSourceFolder = txtFolder + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        di = Directory.CreateDirectory(txtFolder);
                        di = Directory.CreateDirectory(txtSourceFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        di = Directory.CreateDirectory(htmlFolder);
                        di = Directory.CreateDirectory(htmlSourceFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        ZipFile.ExtractToDirectory(file, foldername);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        string[] items = Directory.GetFiles(foldername);
                        foreach (var item in items)
                        {
                            if (Path.GetExtension(item) == ".txt")
                            {
                                File.Move(item, txtSourceFolder + @"\" + item.Substring(item.LastIndexOf(@"\"), (item.Length - item.LastIndexOf(@"\"))));
                                File.Delete(item);
                            }
                            else if (Path.GetExtension(item) == ".html")
                            {
                                File.Move(item, htmlSourceFolder + @"\" + item.Substring(item.LastIndexOf(@"\"), (item.Length - item.LastIndexOf(@"\"))));
                                File.Delete(item);
                            }
                        }
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        FixTXT(txtSourceFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                        ZipFolder(foldername, txtFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                        ZipFolder(foldername, htmlFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                    }
                    else
                    {
                        MessageBox.Show("The specified folder does not contain valid zipped source file.");
                        return;
                    }
                }                                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ZipFolder(string parentFolder, string zipFolder)
        {
            ZipFile.CreateFromDirectory(zipFolder, (parentFolder + @"\" + zipFolder.Substring(zipFolder.LastIndexOf(@"\"), (zipFolder.Length - zipFolder.LastIndexOf(@"\"))) + @".zip"));
            Directory.Delete(zipFolder, true);
        }

        private void FixTXT(string txtSourceFolder)
        {
            string[] files = Directory.GetFiles(txtSourceFolder);
            foreach (string file in files)
            {
                string fileContent = File.ReadAllText(file);
                StringBuilder sb = new StringBuilder(fileContent);

                while (sb.ToString().Contains(" \","))
                {
                    sb.Replace(" \",", "\",");
                }

                string[] spaces = {"            ", "           ", "          ", "         ", "        ", "       ", "      ", "     " };
                foreach (string space in spaces)
                {
                    while (sb.ToString().Contains(space))
                    {
                        sb.Replace(space, " ");
                    }
                }

                sb.Replace(@"\r\n", @"\n");//-----------convert Windows newline to Unix newline
                File.WriteAllText(file, sb.ToString());
            }
        }

        private void Mt_chp_prep_Click(object sender, RoutedEventArgs e)//-----MT CHP Prep section Start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;

                progressBar4.Value = 0;
                progressBar4.Maximum = 4;

                BackgroundWorker bgw8 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw8.WorkerReportsProgress = true;
                bgw8.DoWork += Bgw8_DoWork;
                bgw8.ProgressChanged += Bgw8_ProgressChanged;
                bgw8.RunWorkerCompleted += Bgw8_RunWorkerCompleted;
                bgw8.RunWorkerAsync(PathDisplay4.Text);
                
            }
        }

        private void Bgw8_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw8_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw8_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder containing zipped source file.");
                    return;
                }

                string foldername = pathDisplay + @"\" + "TMS_Source";
                if (Directory.Exists(foldername))
                {
                    MessageBoxButton buttons = MessageBoxButton.YesNo;

                    MessageBoxResult result = MessageBox.Show("Cannot create folder \"TMS_Source\" as it already exists\nDo you wish to overwrite it?", foldername, buttons);
                    if (result == MessageBoxResult.No)
                    {
                        MessageBox.Show("Process Aborted.");
                        return;
                    }
                    else
                    {
                        Directory.Delete(foldername, true);
                    }
                }

                string[] files = Directory.GetFiles(pathDisplay);
                foreach (string file in files)
                {
                    if (Path.GetExtension(file) == ".zip")
                    {
                        DirectoryInfo di = Directory.CreateDirectory(foldername);                        
                        string MtHcSourceFolder = foldername + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));                        
                        (sender as BackgroundWorker).ReportProgress(progress++);
                                                
                        di = Directory.CreateDirectory(MtHcSourceFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                        
                        ZipFile.ExtractToDirectory(file, foldername);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                        string[] items = Directory.GetFiles(foldername);
                        foreach (var item in items)
                        {
                            if (Path.GetExtension(item) == ".txt")
                            {
                                File.Move(item, MtHcSourceFolder + @"\" + item.Substring(item.LastIndexOf(@"\"), (item.Length - item.LastIndexOf(@"\"))));
                                File.Delete(item);
                            }
                            else if (Path.GetExtension(item) == ".html")
                            {
                                ConvertHtmlFileToUTF8bom(item, MtHcSourceFolder);
                               // File.Move(item, MtHcSourceFolder + @"\" + item.Substring(item.LastIndexOf(@"\"), (item.Length - item.LastIndexOf(@"\"))));
                                File.Delete(item);
                            }
                        }
                        (sender as BackgroundWorker).ReportProgress(progress++);                                               
                        
                        ZipFolder(foldername, MtHcSourceFolder);
                        (sender as BackgroundWorker).ReportProgress(progress++);

                    }
                    else
                    {
                        MessageBox.Show("The specified folder does not contain valid zipped source file.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ConvertHtmlFileToUTF8bom(string file, string MtHcSourceFolder)
        {
            using (StreamReader reader = new StreamReader(file))
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(reader.ReadToEnd());
                string fileContent = sb.ToString();
                sb.Clear();
                foreach (char character in fileContent)
                {
                    sb.Append(character);
                }
                sb.Replace(Environment.NewLine, ('\u000A').ToString());//-----------convert Windows newline with Unix newline
                string newFile = MtHcSourceFolder + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\")));                               
                Encoding utfBom = new UTF8Encoding(true, true);//Encoding object will be used as parameter for File.WriteAllText method                           
                File.WriteAllText(newFile, sb.ToString(), utfBom);                
            }
        }

        private void Mt_chp_post_Click(object sender, RoutedEventArgs e)//-----MT CHP Post section Start
        {
            if (PathDisplay4.Text == "")
            {
                MessageBox.Show("Target folder was not selected");
            }
            else
            {
                ToolStack.IsEnabled = false;

                progressBar4.Value = 0;
                progressBar4.Maximum = (Directory.GetFiles(PathDisplay4.Text)).Length;
                
                BackgroundWorker bgw9 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw9.WorkerReportsProgress = true;
                bgw9.DoWork += Bgw9_DoWork;
                bgw9.ProgressChanged += Bgw9_ProgressChanged;
                bgw9.RunWorkerCompleted += Bgw9_RunWorkerCompleted;
                bgw9.RunWorkerAsync(PathDisplay4.Text);
            }
        }

        private void Bgw9_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            ToolStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw9_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar4.Value = e.ProgressPercentage;
        }

        private void Bgw9_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;
            int progress = 0;
            try
            {
                if (Directory.Exists(pathDisplay) == false)
                {
                    MessageBox.Show("The specified path does not lead to valid folder.");
                    return;
                }

                string[] files = Directory.GetFiles(pathDisplay);

                string foldername = pathDisplay + @"\" + "05_Final_Cisco";
                string tempFolder = foldername + @"\" + "Temp";
                if (Directory.Exists(foldername))
                {
                    MessageBoxButton buttons = MessageBoxButton.YesNo;

                    MessageBoxResult result = MessageBox.Show("Cannot create folder \"05_Final_Cisco\" as it already exists\nDo you wish to overwrite it?", foldername, buttons);
                    if (result == MessageBoxResult.No)
                    {
                        MessageBox.Show("Process Aborted.");
                        return;
                    }
                    else
                    {
                        Directory.Delete(foldername, true);
                    }
                }
                DirectoryInfo di = Directory.CreateDirectory(foldername);
                (sender as BackgroundWorker).ReportProgress(progress++);

                foreach (string file in files)
                {
                    if (Path.GetExtension(file) == ".zip")
                    {
                        di = Directory.CreateDirectory(tempFolder);
                        string finalFolder = foldername + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));// + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));
                        string finalFolderForZip = foldername + @"\" + file.Substring(file.LastIndexOf(@"\"), (file.Length - file.LastIndexOf(@"\") - 4));
                        di = Directory.CreateDirectory(finalFolder);

                        ZipFile.ExtractToDirectory(file, tempFolder);                        

                        string[] directories = Directory.GetDirectories(tempFolder, "*", SearchOption.AllDirectories);
                        List<string> engPostFolders = new List<string>();
                        foreach (string directory in directories)
                        {
                            string directoryName = directory.Substring(directory.LastIndexOf(@"\"), (directory.Length - directory.LastIndexOf(@"\")));
                            if (directoryName == @"\ENG_postprocess")
                            {
                                engPostFolders.Add(directory);
                            }
                        }                        

                        foreach ( string engPostFolder in engPostFolders)
                        {
                            int lastSlash = engPostFolder.LastIndexOf(@"\");
                            int secondToLastSlash = lastSlash > 0 ? engPostFolder.LastIndexOf(@"\", lastSlash - 1) : -1;
                            string oldLangCode = engPostFolder.Substring(secondToLastSlash, (engPostFolder.LastIndexOf(@"\") - secondToLastSlash));
                            string langCode = oldLangCode.Replace('-', '_');
                            string finalLangCode;
                            if (langCode == @"\nb_no")
                            {
                                finalLangCode = @"\no_no";
                            }
                            else if (langCode == @"\srl_rs")
                            {
                                finalLangCode = @"\sr_rs";
                            }
                            else
                            {
                                finalLangCode = langCode;
                            }
                            di = Directory.CreateDirectory(finalFolder + @"\" + finalLangCode);
                            string[] contentFiles = Directory.GetFiles(engPostFolder);
                            foreach (string contentFile in contentFiles)
                            {
                                if (Path.GetExtension(contentFile) == ".txt" || Path.GetExtension(contentFile) == ".html")
                                {                                    
                                    File.Move(contentFile, finalFolder + @"\" + finalLangCode + contentFile.Substring(contentFile.LastIndexOf(@"\"), (contentFile.LastIndexOf(@"@") - contentFile.LastIndexOf(@"\"))) + contentFile.Substring(contentFile.LastIndexOf(@"."), (contentFile.Length - contentFile.LastIndexOf(@"."))));
                                }
                            }
                        }                        

                        string[] folders = Directory.GetDirectories(finalFolder);
                        foreach (string folder in folders)
                        {
                            string[] finalFiles = Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
                            foreach (string finalFile in finalFiles)
                            {
                                if (Path.GetExtension(finalFile) == ".txt")
                                {
                                    string fileContentWithQuotes = File.ReadAllText(finalFile);
                                    string fileContent = DeleteQuotes(fileContentWithQuotes);
                                    StringBuilder sb = new StringBuilder(Regex.Replace(fileContent, @"(\d{1,2})-(\d{1,2})-(\d{4})", @"$3-$2-$1"));

                                    while (sb.ToString().Contains(" \","))
                                    {
                                        sb.Replace(" \",", "\",");
                                    }

                                    sb.Replace(Environment.NewLine, ('\u000A').ToString());//-----------convert Windows newline with Unix newline                                    
                                    File.WriteAllText(finalFile, sb.ToString());
                                    sb.Clear();
                                }
                                else if (Path.GetExtension(finalFile) == ".html")
                                {
                                    string fileContent = File.ReadAllText(finalFile);                                    
                                    StringBuilder sb = new StringBuilder(fileContent);

                                    while (sb.ToString().Contains(@"â‹®"))
                                    {
                                        sb.Replace(@"â‹®", @"⋮");
                                    }
                                    while (sb.ToString().Contains(@"â€”"))
                                    {
                                        sb.Replace(@"â€”", @"—");
                                    }

                                    sb.Replace(Environment.NewLine, ('\u000A').ToString());//-----------convert Windows newline with Unix newline
                                    File.WriteAllText(finalFile, sb.ToString());
                                    sb.Clear();
                                }
                            }                           
                        }
                        ZipFile.CreateFromDirectory(finalFolderForZip, finalFolderForZip + @".zip");
                        Directory.Delete(finalFolderForZip, true);
                        Directory.Delete(tempFolder, true);
                        (sender as BackgroundWorker).ReportProgress(progress++);
                    }
                    else
                    {
                        MessageBox.Show("The specified folder does not contain valid zipped source file.");
                        return;
                    }
                }                                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private string DeleteQuotes(string fileContentWithQuotes)
        {
            StringBuilder sbl = new StringBuilder();
            StringBuilder sblTEMP = new StringBuilder();
            using (StringReader reader = new StringReader(fileContentWithQuotes))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains("\"value\":"))
                    {
                        sblTEMP.Append(line);
                        sblTEMP.Replace("\"", "");
                        sblTEMP.Replace("value: ", "\"value\": \"");
                        sblTEMP.Insert(sblTEMP.Length - 1, "\"");
                        line = sblTEMP.ToString();
                        sbl.AppendLine(line);
                        sblTEMP.Clear();
                    }
                    else
                    {
                        if (line == "]")
                        {
                            sbl.Append(line);
                        }
                        else
                        {
                            sbl.AppendLine(line);
                        }                        
                    }
                    
                }
                string result = sbl.ToString();
                sbl.Clear();
                return result;
            }             
        }

        private void StackDropTarget_Drop(object sender, DragEventArgs e)//-----VOP EXCELL STUFF section Start
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] pathT = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (Path.GetExtension(pathT[0]) == ".xlsx")
                {
                    TargetPathDisplay.Text = pathT[0];
                }
                else
                {
                    MessageBox.Show("Incorrect file extension\nPlease use *.xlsx file.");
                }
            }
        }

        private void StackDropSource_Drop(object sender, DragEventArgs e)
        {
            string[] pathS = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (Path.GetExtension(pathS[0]) == ".xlsx")
            {
                SourcePathDisplay.Text = pathS[0];
            }
            else
            {
                MessageBox.Show("Incorrect file extension\nPlease use *.xlsx file.");
            }
        }

        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (TargetPathDisplay.Text == "" || SourcePathDisplay.Text == "")
            {
                MessageBox.Show("Source or/and Target file was not selected");
            }
            else
            {
                VopStack.IsEnabled = false;                
                VopProgressBar.IsIndeterminate = true;
                BackgroundWorker bgw10 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw10.WorkerReportsProgress = true;
                bgw10.DoWork += Bgw10_DoWork;
                bgw10.ProgressChanged += Bgw10_ProgressChanged;
                bgw10.RunWorkerCompleted += Bgw10_RunWorkerCompleted;
                bgw10.RunWorkerAsync(TargetPathDisplay.Text);                
            }
        }

        private void Bgw10_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            VopProgressBar.IsIndeterminate = false;
            VopStack.IsEnabled = true;            
            ArrayList result = (ArrayList)e.Result;
            List<string> languages = (List<string>)result[0];
            int range = (int)result[1];
            RowsCount.Text = range.ToString();
            DropdownLanguages.ItemsSource = languages;
            DropdownLanguages.UpdateLayout();
            MessageBox.Show("Done");
        }

        private void Bgw10_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            VopProgressBar.Value = e.ProgressPercentage;
        }

        private void Bgw10_DoWork(object sender, DoWorkEventArgs e)
        {
            string pathDisplay = (string)e.Argument;            
            try
            {
                //Create excel instance
                Excel.Application excel = new Excel.Application();
                //Create target workbook
                Excel.Workbook workbook = excel.Workbooks.Open(pathDisplay);
                //Open the source sheet
                Excel.Worksheet sourceSheet = workbook.Worksheets[1];
                //Get the used Range
                Excel.Range usedRange = sourceSheet.UsedRange;
                int range = usedRange.Rows.Count;
                List<string> languages = new List<string>();
                for (int i = 2; i <= range; i++)
                {
                    string language = (string)(sourceSheet.Cells[i, 2] as Excel.Range).Value;
                    bool alreadyOnTheList = false;
                    foreach (string lang in languages)
                    {                        
                        if (lang == language)
                        {
                            alreadyOnTheList = true;
                        }
                    }
                    if (alreadyOnTheList == false)
                    {
                        languages.Add(language);                        
                    }
                }
                //Close workbook                    
                workbook.Close(false);
                //Kill excelapp
                excel.Quit();
                var result = new ArrayList() { languages, range};
                e.Result = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }      
               
        private void CopyTrans_Click(object sender, RoutedEventArgs e)// ----- Copy Translation BUTTON
        {
            if (TargetPathDisplay.Text == "" || SourcePathDisplay.Text == "")
            {
                MessageBox.Show("Source or Target file was not selected");
            }
            else if (DropdownLanguages.SelectedItem as string == null)
            {
                MessageBox.Show("Language is not selected");
            }
            else
            {
                VopStack.IsEnabled = false;
                VopProgressBar.Value = 0;
                VopProgressBar.Maximum = int.Parse(RowsCount.Text);
                string[] bgwBundle = { TargetPathDisplay.Text, SourcePathDisplay.Text, DropdownLanguages.SelectedItem as string };
                BackgroundWorker bgw11 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw11.WorkerReportsProgress = true;
                bgw11.DoWork += Bgw11_DoWork;
                bgw11.ProgressChanged += Bgw11_ProgressChanged;
                bgw11.RunWorkerCompleted += Bgw11_RunWorkerCompleted;
                bgw11.RunWorkerAsync(bgwBundle);

            }
        }

        private void Bgw11_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //-------------------------------> enable wpf controls
            VopStack.IsEnabled = true;
            MessageBox.Show("Done");
        }

        private void Bgw11_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            VopProgressBar.Value = e.ProgressPercentage;
        }

        private void Bgw11_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] bgwBundle = (string[])e.Argument;
            string target = bgwBundle[0];
            string source = bgwBundle[1];
            string language = bgwBundle[2];
            int progress = 0;
            try
            {
                Excel.Application excel = new Excel.Application();
                //Create target workbook
                Excel.Workbook excelTarget = excel.Workbooks.Open(target);
                //Create source workbook
                Excel.Workbook excelSource = excel.Workbooks.Open(source);

                //Open the source sheet
                Excel.Worksheet sourceSheet = excelSource.Worksheets[1];
                //Open the target sheet
                Excel.Worksheet targetSheet = excelTarget.Worksheets[1];

                //Get the used Range
                Excel.Range usedRange = sourceSheet.UsedRange;
                int range = usedRange.Rows.Count;

                for (int i = 2; i <= range; i++)
                {
                    if ((string)(sourceSheet.Cells[i, 2] as Excel.Range).Value == language)
                    {
                        targetSheet.Cells[i, 23] = (string)(sourceSheet.Cells[i, 23] as Excel.Range).Value;                        
                    }
                    (sender as BackgroundWorker).ReportProgress(progress++);
                }
                
                //Close and save target workbooks
                excelTarget.Close(true);
                excelSource.Close(true);
                //Kill excelapp
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }
    }
}
