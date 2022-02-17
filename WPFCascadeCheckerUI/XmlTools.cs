using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Windows;

namespace WPFCascadeCheckerUI
{
    class XmlTools
    {        
        public static void CheckXml(string file, StreamWriter sw, string uicOverride)
        {
            if (Path.GetExtension(file) == ".xml" || Path.GetExtension(file) == ".ditamap" || Path.GetExtension(file) == ".dita")
            {
                string filePath = "";
                try
                {
                    XmlDocument xml = new XmlDocument();
                                       
                    using (XmlReader xr = new XmlTextReader(file) { Namespaces = false })
                    {                        
                        xml = new PositionXmlDocument();
                        xml.Load(xr);
                    }

                    var root = xml.SelectNodes("//*");

                    foreach (var element in root)
                    {
                        var node = (PositionXmlElement)element;
                        if (node.Name == "menucascade")//---------------------------------------------MENUCASCADE SECTION
                        {
                            //CheckIfMenucascadeHaveText(node);
                            if (node.Name == "menucascade" && node.InnerText != "")
                            {
                                //get all child nodes of type XmlText                        
                                var childNodes = node.ChildNodes.OfType<XmlText>();
                                if (childNodes.Count() != 0)
                                {
                                    if (filePath != file)
                                    {
                                        sw.WriteLine("\n\n\nFile: {0}", file);
                                        filePath = file;
                                    }
                                    sw.WriteLine("\n  **Forbidden text element inside node: \"{0}\"! Line: {1}, Position: {2}\n        Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                    int xtCounter = 1;
                                    foreach (var xT in childNodes)
                                    {
                                        sw.WriteLine("            Forbidden text element #{0}: {1}", xtCounter, xT.OuterXml);
                                        xtCounter++;
                                    }
                                }
                            }
                            if (node.Name == "menucascade" && node.IsEmpty)
                            {
                                if (filePath != file)
                                {
                                    sw.WriteLine("\n\n\nFile: {0}", file);
                                    filePath = file;
                                }
                                sw.WriteLine("\n  **Menucascade node is empty: \"{0}\"! Line: {1}, Position: {2}\n        Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                            }
                        }
                        if (node.Name == "uicontrol")//-----------------------------------------------UICONTROL SECTION
                        {
                            //Check if uicontrol has elements other than those listed.
                            if (node.HasChildNodes)
                            {
                                if (uicOverride == "standard")
                                {
                                    var xmlText = node.ChildNodes.OfType<XmlText>();
                                    var keyWords = node.GetElementsByTagName("keyword");
                                    var varnames = node.GetElementsByTagName("varname");
                                    var images = node.GetElementsByTagName("image");
                                    var apiNames = node.GetElementsByTagName("apiname");
                                    var winTitles = node.GetElementsByTagName("wintitle");
                                    var cmdNames = node.GetElementsByTagName("cmdname");
                                    if (node.ChildNodes.Count > (xmlText.Count() + keyWords.Count + varnames.Count + images.Count + winTitles.Count + apiNames.Count + cmdNames.Count))
                                    {
                                        if (filePath != file)
                                        {
                                            sw.WriteLine("\n\n\nFile: {0}", file);
                                            filePath = file;
                                        }
                                        sw.WriteLine("\n  **Node: \"{0}\" have forbidden child nodes!  Line: {1}, Position: {2}\n      Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                        int uiCounter = 1;
                                        foreach (var child in node.ChildNodes)
                                        {
                                            var uiChild = (XmlNode)child;
                                            if (uiChild.Name != "keyword" && uiChild.Name != "varname" && uiChild.Name != "image" && uiChild.Name != "apiname" && uiChild.Name != "wintitle" && uiChild.Name != "cmdname" && uiChild.NodeType.ToString() != "Text")
                                            {
                                                sw.WriteLine("            Forbidden element #{0}: {1}", uiCounter, uiChild.OuterXml);
                                                uiCounter++;
                                            }
                                        }
                                    }
                                }
                                else if (uicOverride == "simplified")
                                {                                    
                                    var uicontrols = node.GetElementsByTagName("uicontrol");
                                    var menucascades = node.GetElementsByTagName("menucascade");                                    
                                    if ((uicontrols.Count + menucascades.Count) > 0)
                                    {
                                        if (filePath != file)
                                        {
                                            sw.WriteLine("\n\n\nFile: {0}", file);
                                            filePath = file;
                                        }
                                        sw.WriteLine("\n  **Node: \"{0}\" have forbidden child nodes!  Line: {1}, Position: {2}\n      Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                        int uiCounter = 1;
                                        foreach (var child in node.ChildNodes)
                                        {
                                            var uiChild = (XmlNode)child;
                                            if (uiChild.Name != "keyword" && uiChild.Name != "varname" && uiChild.Name != "image" && uiChild.NodeType.ToString() != "Text")
                                            {
                                                sw.WriteLine("            Forbidden element #{0}: {1}", uiCounter, uiChild.OuterXml);
                                                uiCounter++;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Application encountered error while transferring checkbox settings");
                                    return;
                                }                                
                            }
                        }
                        if (node.Name == "cite")//----------------------------------------------------CITE SECTION
                        {
                            if (node.HasChildNodes)
                            {
                                XmlNodeList citeChilds = node.GetElementsByTagName("xref");
                                if (citeChilds.Count != 0)
                                {
                                    if (filePath != file)
                                    {
                                        sw.WriteLine("\n\n\nFile: {0}", file);
                                        filePath = file;
                                    }
                                    sw.WriteLine("\n  **Node: \"{0}\" have forbidden child node: \"xref\"!  Line: {1}, Position: {2}\n      Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                }

                            }
                        }
                        if (node.Name == "image" && node.InnerText != "")//----------------------------------------------------IMAGE SECTION
                        {
                            var imageText = node.ChildNodes.OfType<XmlText>();
                            if (imageText.Count() != 0)
                            {
                                if (filePath != file)
                                {
                                    sw.WriteLine("\n\n\nFile: {0}", file);
                                    filePath = file;
                                }
                                sw.WriteLine("\n  **Forbidden text element inside node: \"{0}\"! Line: {1}, Position: {2}\n        Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                int itCounter = 1;
                                foreach (var iT in imageText)
                                {
                                    sw.WriteLine("            Forbidden text element #{0}: {1}", itCounter, iT.OuterXml);
                                    itCounter++;
                                }
                            }
                        }
                        if (node.Name == "keyword")//------------------------------------------------KEYWORD SECTION
                        {
                            //Check if keyword has elements other than those listed.
                            if (node.HasChildNodes)
                            {                                                                
                                var xmlText = node.ChildNodes.OfType<XmlText>();
                                var comments = node.ChildNodes.OfType<XmlComment>();
                                var cleanups = node.GetElementsByTagName("required-cleanup");
                                var tms = node.GetElementsByTagName("tm");

                                if (node.ChildNodes.Count > (xmlText.Count() + comments.Count() + cleanups.Count + tms.Count))
                                {
                                    if (filePath != file)
                                    {
                                        sw.WriteLine("\n\n\nFile: {0}", file);
                                        filePath = file;
                                    }
                                    sw.WriteLine("\n  **Node: \"{0}\" have forbidden child nodes!  Line: {1}, Position: {2}\n      Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                    int keyWordCounter = 1;
                                    foreach (var child in node.ChildNodes)
                                    {
                                        var keyWordChild = (XmlNode)child;
                                        if (keyWordChild.Name != "required-cleanup" && keyWordChild.Name != "tm" && keyWordChild.NodeType.ToString() != "Comment" && keyWordChild.NodeType.ToString() != "Text")
                                        {
                                            sw.WriteLine("            Forbidden element #{0}: {1}", keyWordCounter, keyWordChild.OuterXml);
                                            keyWordCounter++;
                                        }
                                    }
                                }                                                            
                            }
                        }
                        if (node.Name == "cmd")//------------------------------------------------CMD SECTION
                        {
                            if (node.HasChildNodes)
                            {
                                var kwds = node.GetElementsByTagName("kwd");
                                var synphs = node.GetElementsByTagName("synph");
                                int kwdCounter = 0;
                                foreach (PositionXmlElement synph in synphs)
                                {
                                    var nestedKWDs = synph.GetElementsByTagName("kwd");
                                    kwdCounter += nestedKWDs.Count;
                                }
                                if (kwds.Count > 0 && kwdCounter != kwds.Count)
                                {
                                    if (filePath != file)
                                    {
                                        sw.WriteLine("\n\n\nFile: {0}", file);
                                        filePath = file;
                                    }
                                    sw.WriteLine("\n  **Forbidden text element inside node: \"{0}\"! Line: {1}, Position: {2}\n        Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                    int kwdIssueCounter = 1;
                                    foreach (var child in node.ChildNodes)
                                    {
                                        var kwdChild = (XmlNode)child;
                                        if (kwdChild.Name == "kwd")
                                        {
                                            sw.WriteLine("            Forbidden element #{0}: {1}       *this element is allowed in node <cmd> only inside of <synph></synph> parent structure", kwdIssueCounter, kwdChild.OuterXml);
                                            kwdIssueCounter++;
                                        }
                                    }
                                }
                            }
                        }
                        if (node.Name == "sep")//----------------------------------------------------SEP SECTION
                        {
                            if (node.HasChildNodes)
                            {
                                XmlNodeList citeChilds = node.GetElementsByTagName("var");
                                if (citeChilds.Count != 0)
                                {
                                    if (filePath != file)
                                    {
                                        sw.WriteLine("\n\n\nFile: {0}", file);
                                        filePath = file;
                                    }
                                    sw.WriteLine("\n  **Node: \"{0}\" have forbidden child node: \"var\"!  Line: {1}, Position: {2}\n      Affected node: {3}", node.Name, node.LineNumber, node.LinePosition, node.OuterXml);
                                }

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    if (filePath != file)
                    {
                        sw.WriteLine("\n\n\nFile: {0}", file);
                        filePath = file;
                    }
                    sw.WriteLine("    {0}\n", e.Message);
                }
            }            
        }

        internal static void CreateConfigFile(string directory)
        {
            // this method generates config file for TMS batch, using folder with prepared subfolder structure (name of each subfolder must be language code in TMS format)
            if (Directory.Exists(directory))
            {
                string[] languageFolders = Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly);
                XDocument config = new XDocument(
                new XDeclaration("1.0", "utf-16", "yes"),
                new XElement("BatchUpload")
                );

                foreach (var folder in languageFolders)
                {
                    string languageCode = folder.Substring((folder.LastIndexOf(@"\") + 1), (folder.Length - folder.LastIndexOf(@"\") - 1));
                    string[] fullFilePath = Directory.GetFiles(folder, "*", SearchOption.TopDirectoryOnly);
                    string filePath = fullFilePath[0].Substring(fullFilePath[0].LastIndexOf(languageCode), (fullFilePath[0].Length - fullFilePath[0].LastIndexOf(languageCode)));

                    //Add newchold nodes to the <BatchUpload> node
                    config.Element("BatchUpload").Add(new XElement("Job", new XAttribute("Source", "en-us"), new XAttribute("Target", languageCode),
                        new XElement("File", new XAttribute("Fullpath", filePath))));
                }

                config.Save(directory + @"\format.config");
            }
        }

        internal static void CheckForXmlLang(string file, StreamWriter sw)
        {
            if (Path.GetExtension(file) == ".xml" || Path.GetExtension(file) == ".ditamap" || Path.GetExtension(file) == ".dita")
            {
                string filePath = "";
                try
                {
                    XmlDocument xml = new XmlDocument();

                    using (XmlReader xr = new XmlTextReader(file) { Namespaces = false })
                    {
                        xml = new PositionXmlDocument();
                        xml.Load(xr);
                    }

                    //var root = xml.SelectNodes("//*");
                    var root = xml.DocumentElement;
                    string missingXmlLangPath = file.Substring(0, file.LastIndexOf('\\'));
                    if (root.HasAttribute("xml:lang") == false)
                    {                        
                        sw.WriteLine("\n\n{0} - incorrect", missingXmlLangPath);
                    }
                }
                catch (Exception e)
                {
                    if (filePath != file)
                    {
                        sw.WriteLine("\n\n\nFile: {0}", file);
                        filePath = file;
                    }
                    sw.WriteLine("    {0}\n", e.Message);
                }
            }
        }
    }
}
