using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SheetDock
{
    /// <summary>
    /// Stores the virtual folder structure inside the workbook via CustomXMLParts,
    /// without modifying the actual sheets.
    /// </summary>
    public static class FolderStore
    {
        private const string NamespaceUri = "urn:sheetdock:folders:v1";
        private const string RootTag = "SheetDock";

        public static List<FolderInfo> Load(Excel.Workbook workbook)
        {
            var result = new List<FolderInfo>();
            if (workbook == null) return result;

            var part = FindPart(workbook.CustomXMLParts);
            if (part == null) return result;

            try
            {
                var doc = XDocument.Parse(part.XML);
                var root = doc.Root;
                if (root == null) return result;

                foreach (var folderEl in root.Elements("Folder"))
                {
                    var folder = new FolderInfo
                    {
                        Name = (string)folderEl.Attribute("name") ?? "Folder"
                    };

                    foreach (var sheetEl in folderEl.Elements("Sheet"))
                    {
                        var name = (string)sheetEl.Attribute("name");
                        if (!string.IsNullOrEmpty(name))
                            folder.Sheets.Add(name);
                    }

                    result.Add(folder);
                }
            }
            catch
            {
                // Corrupt XML -> return an empty list.
            }

            return result;
        }

        public static void Save(Excel.Workbook workbook, IEnumerable<FolderInfo> folders)
        {
            if (workbook == null) return;

            var root = new XElement(RootTag,
                new XAttribute("xmlns", NamespaceUri),
                folders.Select(f =>
                    new XElement("Folder",
                        new XAttribute("name", f.Name),
                        f.Sheets.Select(s =>
                            new XElement("Sheet", new XAttribute("name", s))))));

            var xml = new XDocument(root).ToString();

            var parts = workbook.CustomXMLParts;
            var existing = FindPart(parts);

            if (existing != null)
            {
                existing.LoadXML(xml);
            }
            else
            {
                parts.Add(xml);
            }
        }

        private static Office.CustomXMLPart FindPart(Office.CustomXMLParts parts)
        {
            if (parts == null) return null;

            foreach (Office.CustomXMLPart p in parts)
            {
                try
                {
                    var doc = XDocument.Parse(p.XML);
                    var root = doc.Root;
                    if (root != null &&
                        root.Name.LocalName == RootTag &&
                        root.Name.NamespaceName == NamespaceUri)
                    {
                        return p;
                    }
                }
                catch { }
            }

            return null;
        }
    }

    public class FolderInfo
    {
        public string Name { get; set; } = "Folder";
        public List<string> Sheets { get; } = new List<string>();
    }
}