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
        internal static void CreateLanguageFolders(string filePath, bool passoloMarker)
        {            
            try
            {
                //Create array of filenames with path from directory and subfolders
                string[] files = Directory.GetFiles(filePath, "*", SearchOption.AllDirectories);

                foreach (var file in files)
                {
                    if (Path.GetExtension(file) == ".sdlxliff") //only sdlxliff files exported from pasolo are processed
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
                XmlTools.CreateConfigFile(filePath, passoloMarker);
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

        public static string CheckLanguageCodeTMS(string languageCode)
        {            
            var tmsLangCodesDic = new Dictionary<string, string>(){
                {"arsa", "ar-sa"},
                {"bgbg", "bg-bg"},
                {"caes", "ca-es"},
                {"cscz", "cs-cz"},
                {"dadk", "da-dk"},
                {"dede", "de-de"},
                {"elgr", "el-gr"},
                {"esco", "es-co"},
                {"eses", "es-es"},
                {"etee", "et-ee"},
                {"fifi", "fi-fi"},
                {"frca", "fr-ca"},
                {"frfr", "fr-fr"},
                {"heil", "he-il"},
                {"hrhr", "hr-hr"},
                {"huhu", "hu-hu"},
                {"itit", "it-it"},
                {"jajp", "ja-jp"},
                {"kokr", "ko-kr"},
                {"ltlt", "lt-lt"},
                {"lvlv", "lv-lv"},
                {"nlnl", "nl-nl"},
                {"nono", "nb-no"},
                {"plpl", "pl-pl"},
                {"ptbr", "pt-br"},
                {"ptpt", "pt-pt"},
                {"roro", "ro-ro"},
                {"ruru", "ru-ru"},
                {"sksk", "sk-sk"},
                {"slsi", "sl-si"},
                {"srrs", "srl-rs"},
                {"svse", "sv-se"},
                {"thth", "th-th"},
                {"trtr", "tr-tr"},
                {"ukua", "uk-ua"},
                {"zhcn", "zh-cn"},
                {"zhhk", "zh-hk"},
                {"zhtw", "zh-tw"}
            };

            try
            {
                string result = tmsLangCodesDic[languageCode];
                return result;
            }
            catch (KeyNotFoundException)
            {
                MessageBox.Show("Unsupported language: " + languageCode);
                string result = null;
                return result;
            }                    
        }

        public static string GetNumericLanguageCode(string languageCode)
        {
            string finalLanguageCode = (languageCode.Replace('-', '_')).ToLower();
            var numLangCodesDic = new Dictionary<string, string>(){
                {"ar_sa", "9 1\t1 1"},
                {"bg_bg", "9 1\t2 1"},
                {"ca_es", "9 1\t3 1"},
                {"cs_cz", "9 1\t5 1"},
                {"da_dk", "9 1\t6 1"},
                {"de_de", "9 1\t7 1"},
                {"el_gr", "9 1\t8 1"},
                {"es_es", "9 1\t10 3"},
                {"es_xl", "9 1\t10 9"},
                {"et_ee", "9 1\t37 1"},
                {"fi_fi", "9 1\t11 1"},
                {"fr_ca", "9 1\t12 3"},
                {"fr_fr", "9 1\t12 1"},
                {"he_il", "9 1\t13 1"},
                {"hr_hr", "9 1\t26 1"},
                {"hu_hu", "9 1\t14 1"},
                {"it_it", "9 1\t16 1"},
                {"ja_jp", "9 1\t17 1"},
                {"ko_kr", "9 1\t18 1"},
                {"lt_lt", "9 1\t39 1"},
                {"lv_lv", "9 1\t38 1"},
                {"nl_nl", "9 1\t19 1"},
                {"nb_no", "9 1\t20 1"},
                {"pl_pl", "9 1\t21 1"},
                {"pt_br", "9 1\t22 1"},
                {"pt_pt", "9 1\t22 2"},
                {"ro_ro", "9 1\t24 1"},
                {"ru_ru", "9 1\t25 1"},
                {"sk_sk", "9 1\t27 1"},
                {"sl_si", "9 1\t36 1"},
                {"sr_rs", "9 1\t26 2"},
                {"sv_se", "9 1\t29 1"},
                {"th_th", "9 1\t30 1"},
                {"tr_tr", "9 1\t31 1"},
                {"uk_ua", "9 1\t34 1"},
                {"zh_cn", "9 1\t4 2"},
                {"zh_hk", "9 1\t4 3"},
                {"zh_tw", "9 1\t4 1"},
                {"en_gb", "9 1\t9 2"},
                {"id_id", "9 1\t33 1"}
            };

            string result = numLangCodesDic[finalLanguageCode];

            return result;
        }

        public static string ConvertGloLanguageCode(string languageCode)
        {
            string finalLanguageCode = (languageCode.Replace('-', '_')).ToLower();
            var gloLangCodesDic = new Dictionary<string, string>(){
                {"ar_sa", "ar-SA"},
                {"bg_bg", "bg-BG"},
                {"ca_es", "ca-ES"},
                {"cs_cz", "cs-CZ"},
                {"da_dk", "da-DK"},
                {"de_de", "de-DE"},
                {"el_gr", "el-GR"},
                {"es_es", "es-ES"},
                {"es_xl", "es-CO"},
                {"et_ee", "et-EE"},
                {"fi_fi", "fi-FI"},
                {"fr_ca", "fr-CA"},
                {"fr_fr", "fr-FR"},
                {"he_il", "he-IL"},
                {"hr_hr", "hr-HR"},
                {"hu_hu", "hu-HU"},
                {"it_it", "it-IT"},
                {"ja_jp", "ja-JP"},
                {"ko_kr", "ko-KR"},
                {"lt_lt", "lt-LT"},
                {"lv_lv", "lv-LV"},
                {"nl_nl", "nl-NL"},
                {"nb_no", "nb-NO"},
                {"pl_pl", "pl-PL"},
                {"pt_br", "pt-BR"},
                {"pt_pt", "pt-PT"},
                {"ro_ro", "ro-RO"},
                {"ru_ru", "ru-RU"},
                {"sk_sk", "sk-SK"},
                {"sl_si", "sl-SI"},
                {"sr_rs", "sr-Latn-CS"},
                {"sv_se", "sv-SE"},
                {"th_th", "th-TH"},
                {"tr_tr", "tr-TR"},
                {"uk_ua", "uk-UA"},
                {"zh_cn", "zh-CN"},
                {"zh_hk", "zh-HK"},
                {"zh_tw", "zh-TW"},
                {"en_gb", "en-GB"},
                {"id_id", "id-ID"}
            };

            string result = gloLangCodesDic[finalLanguageCode];

            return result;
        }
    }
}
