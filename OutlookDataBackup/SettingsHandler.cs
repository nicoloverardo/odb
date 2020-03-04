using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace OutlookDataBackup
{
    public class SettingsHandler
    {
        private string SettingsFolder { get; set; } 
        private string ConfigFilePath { get; set; }
        public string Token { get; set; }
        public string ConflictBehavior { get; set; }
        public string SplitSize { get; set; }
        public Dictionary<PstFile, List<PstFile>> Files { get; set; }

        public SettingsHandler(string settingsFolder)
        {
            if (!Directory.Exists(settingsFolder))
            {
                Directory.CreateDirectory(settingsFolder);
            }

            SettingsFolder = settingsFolder;
            ConfigFilePath = SettingsFolder + @"\settings";

            if (!File.Exists(settingsFolder + @"\settings"))
            {
                CreateXml();
            }     

            Load();
        }

        private void CreateXml()
        {
            using (XmlWriter writer = XmlWriter.Create(ConfigFilePath))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Settings");
                writer.WriteElementString("Token", "");
                writer.WriteElementString("ConflictBehavior", "0");
                writer.WriteElementString("SplitSize", "-v2g");
                writer.WriteStartElement("Files");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }

        public void Load()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(ConfigFilePath);

            Token = Unprotect(xmlDoc.DocumentElement.SelectSingleNode("/Settings/Token").Value); // OR INNERTEXT?
            ConflictBehavior = xmlDoc.DocumentElement.SelectSingleNode("/Settings/ConflictBehavior").Value == "0" ?
                "replace" :
                "rename";
            SplitSize = xmlDoc.DocumentElement.SelectSingleNode("/Settings/SplitSize").Value;
            Files = ReadFilesSetting(xmlDoc);

            xmlDoc.Save(ConfigFilePath);
            xmlDoc = null;
            GC.Collect();
        }

        public Dictionary<PstFile, List<PstFile>> ReadFilesSetting(XmlDocument xmlDoc = null)
        {
            if (xmlDoc == null)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(ConfigFilePath);
            }

            XmlNode files = xmlDoc.DocumentElement.SelectSingleNode("/Settings/Files");
            Dictionary<PstFile, List<PstFile>> finalList = new Dictionary<PstFile, List<PstFile>>();

            if (files.ChildNodes.Count == 0) return finalList;

            foreach (XmlNode file in files.ChildNodes)
            {
                string name = file.Attributes["name"].Value;
                string location = file.ChildNodes.Item(0).Value;
                string destination = file.ChildNodes.Item(1).Value;
                long size = Convert.ToInt64(file.ChildNodes.Item(2).Value);
                double progress = Convert.ToDouble(file.ChildNodes.Item(3).Value);

                List<PstFile> zipParts = new List<PstFile>();
                if (file.LastChild.Name == "ZipParts")
                {
                    foreach (XmlNode part in file.LastChild.ChildNodes)
                    {
                        string partName = part.Attributes["name"].Value;
                        string partLocation = part.ChildNodes.Item(0).Value;
                        string partDestination = file.ChildNodes.Item(1).Value;
                        long partSize = Convert.ToInt64(file.ChildNodes.Item(2).Value);
                        double partProgress = Convert.ToDouble(file.ChildNodes.Item(3).Value);
                        int partHash = Convert.ToInt32(file.ChildNodes.Item(4).Value);

                        zipParts.Add(new PstFile(partName, partLocation, partDestination, partSize, partProgress, partHash));
                    }
                }

                finalList.Add(new PstFile(name, location, destination, size, progress), zipParts);
            }

            return finalList;
        }

        public void SaveFilesSetting(XmlDocument xmlDoc = null)
        {
            if (xmlDoc == null)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(ConfigFilePath);
            }

            XmlNode filesNode = xmlDoc.DocumentElement.SelectSingleNode("/Settings/Files");
            foreach (var item in Files)
            {
                XmlElement itemNode = xmlDoc.CreateElement("File");
                XmlAttribute attribute = xmlDoc.CreateAttribute("name");
                attribute.Value = item.Key.Name;
                itemNode.Attributes.Append(attribute);

                XmlElement locationNode = xmlDoc.CreateElement("Location");
                locationNode.Value = item.Key.Path;
                itemNode.AppendChild(locationNode);

                XmlElement destinationNode = xmlDoc.CreateElement("Destination");
                destinationNode.Value = item.Key.Destination;
                itemNode.AppendChild(destinationNode);

                XmlElement sizeNode = xmlDoc.CreateElement("Size");
                sizeNode.Value = item.Key.Length.ToString();
                itemNode.AppendChild(sizeNode);

                XmlElement progressNode = xmlDoc.CreateElement("Progress");
                progressNode.Value = item.Key.Progress.ToString();
                itemNode.AppendChild(progressNode);

                if (item.Value.Count != 0)
                {
                    XmlElement parts = xmlDoc.CreateElement("ZipParts");
                    foreach (var part in item.Value)
                    {
                        XmlElement partNode = xmlDoc.CreateElement("Part");
                        XmlAttribute partAttribute = xmlDoc.CreateAttribute("name");
                        partAttribute.Value = part.Name;
                        partNode.Attributes.Append(partAttribute);

                        XmlElement partLocationNode = xmlDoc.CreateElement("Location");
                        partLocationNode.Value = part.Path;
                        partNode.AppendChild(partLocationNode);

                        XmlElement partDestinationNode = xmlDoc.CreateElement("Destination");
                        partDestinationNode.Value = part.Destination;
                        partNode.AppendChild(partDestinationNode);

                        XmlElement partSizeNode = xmlDoc.CreateElement("Size");
                        partSizeNode.Value = part.Length.ToString();
                        partNode.AppendChild(partSizeNode);

                        XmlElement partProgressNode = xmlDoc.CreateElement("Progress");
                        partProgressNode.Value = part.Progress.ToString();
                        partNode.AppendChild(partProgressNode);

                        XmlElement partHashNode = xmlDoc.CreateElement("Hash");
                        partHashNode.Value = (new FileInfo(part.Path)).GetHashCode().ToString();
                        partNode.AppendChild(partHashNode);

                        parts.AppendChild(partNode);
                    }

                    itemNode.AppendChild(parts);
                }

                filesNode.AppendChild(itemNode);
            }

            xmlDoc.Save(ConfigFilePath);
        }

        public void Clear(XmlDocument xmlDoc = null)
        {
            if (xmlDoc == null)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(ConfigFilePath);
            }

            xmlDoc.DocumentElement.SelectSingleNode("/Settings/Token").Value = "";
            xmlDoc.DocumentElement.SelectSingleNode("/Settings/ConflictBehavior").Value = "0";
            xmlDoc.DocumentElement.SelectSingleNode("/Settings/SplitSize").Value = "";
            xmlDoc.DocumentElement.SelectSingleNode("/Settings/Files").RemoveAll();

            xmlDoc.Save(ConfigFilePath);
            xmlDoc = null;
            GC.Collect();
        }

        private string LoadSingleSetting(string setting, XmlDocument xmlDoc = null)
        {
            if (xmlDoc == null)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(ConfigFilePath);
            }

            string result = xmlDoc.DocumentElement.SelectSingleNode("/Settings/" + setting).Value;

            xmlDoc.Save(ConfigFilePath);
            xmlDoc = null;

            return result;
        }

        private void SaveSingleSetting(string setting, string value, XmlDocument xmlDoc = null)
        {
            if (xmlDoc == null)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.Load(ConfigFilePath);
            }

            xmlDoc.DocumentElement.SelectSingleNode("/Settings/" + setting).Value = value;

            xmlDoc.Save(ConfigFilePath);
            xmlDoc = null;
            GC.Collect();
        }

        public void Save()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(ConfigFilePath);

            xmlDoc.DocumentElement.SelectSingleNode("/Settings/Token").Value = Protect(Token);
            xmlDoc.DocumentElement.SelectSingleNode("/Settings/ConflictBehavior").Value = ConflictBehavior == "0" ?
                "replace" :
                "rename";

            xmlDoc.DocumentElement.SelectSingleNode("/Settings/SplitSize").Value = SplitSize;
            SaveFilesSetting(xmlDoc);

            xmlDoc.Save(ConfigFilePath);
            xmlDoc = null;
            GC.Collect();
        }

        private static string Protect(string str)
        {
            byte[] entropy = Encoding.UTF8.GetBytes(Assembly.GetExecutingAssembly().FullName);
            byte[] data = Encoding.UTF8.GetBytes(str);
            string protectedData = Convert.ToBase64String(ProtectedData.Protect(data, entropy, DataProtectionScope.CurrentUser));
            return protectedData;
        }

        private static string Unprotect(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";

            byte[] protectedData = Convert.FromBase64String(str);
            byte[] entropy = Encoding.UTF8.GetBytes(Assembly.GetExecutingAssembly().FullName);
            string data = Encoding.UTF8.GetString(ProtectedData.Unprotect(protectedData, entropy, DataProtectionScope.CurrentUser));
            return data;
        }
    }
}