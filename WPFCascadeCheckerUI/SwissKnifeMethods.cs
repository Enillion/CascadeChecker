using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.IO.Compression;

namespace WPFCascadeCheckerUI
{
    class SwissKnifeMethods
    {
        internal static void CreateLanguageFolders(string filePath)
        {            
            try
            {
                //Create array of filenames with path from directory and subfolders
                string[] files = Directory.GetFiles(filePath, "*", SearchOption.AllDirectories);

                foreach (var file in files)
                {
                    if (Path.GetExtension(file) == ".sdlxliff") //only sdlxliff files exported frpm pasolo are processed
                    {
                        string languageCode = file.Substring((file.LastIndexOf("_") + 1), ((file.Length - 10) - file.LastIndexOf("_"))).ToLower();
                        languageCode = CheckLanguageCode(languageCode);
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
            }
            catch (Exception zipex)
            {
                MessageBox.Show(zipex.Message);
            }
            MessageBox.Show("Done");
        }

        private static void CreateConfigFile(object filepath)
        {
            
        }

        public static string CheckLanguageCode(string languageCode)
        {
            string result = languageCode;
            string[] passoloLanguageCodes = {"hy-am", "az-latn-az", "bs-latn-ba", "my-mm", "ka-ge", "ga-ie", "kn-in", "kk-kz", "km-kh", "lo-la", "mt-mt",
                "mr-in", "sr-cyrl-rs", "sr-latn-rs", "sr-latn-cs", "es-419", "uz-latn-uz", "cy-gb"};

            string[] tmsLanguageCodes = { "hy", "az", "bs", "my", "ka", "ga", "kn", "kk", "km", "lo", "mt", "mr", "sr-rs", "srl-rs", "srl-rs", "es-xl", "uzl", "cy" };

            for (int i = 0; i < passoloLanguageCodes.Length; i++)
            {
                string xxx = passoloLanguageCodes[i];
                if (result == xxx)
                {
                    result = tmsLanguageCodes[i];
                    break;
                }
            }
            return result;
        }
    }
}
