using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Docx2Html;

public static class UriFixer
{
    public static void FixInvalidUri(Stream fs, Func<string, string> invalidUriHandler)
    {

        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        using ZipArchive? za = new ZipArchive(fs, ZipArchiveMode.Update);
        foreach (ZipArchiveEntry? entry in za.Entries.ToList())
        {
            if (!entry.Name.EndsWith(".rels"))
                continue;
            bool replaceEntry = false;
            XDocument? entryXDoc;
            using (Stream? entryStream = entry.Open())
            {
                try
                {
                    entryXDoc = XDocument.Load(entryStream);
                    if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                    {
                        IEnumerable<XElement>? urisToCheck = entryXDoc
                            .Descendants(relNs + "Relationship")
                            .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                        foreach (XElement? rel in urisToCheck)
                        {
                            string? target = (string)rel.Attribute("Target");
                            if (target != null)
                            {
                                try
                                {
                                    _ = new Uri(target);
                                }
                                catch (UriFormatException)
                                {
                                    string? newUri = invalidUriHandler(target);
                                    rel.Attribute("Target")!.Value = newUri;
                                    replaceEntry = true;
                                }
                            }
                        }
                    }
                }
                catch (XmlException)
                {
                    continue;
                }
            }
            if (replaceEntry)
            {
                string? fullName = entry.FullName;
                entry.Delete();
                ZipArchiveEntry? newEntry = za.CreateEntry(fullName);
                using StreamWriter? writer = new StreamWriter(newEntry.Open());
                using XmlWriter? xmlWriter = XmlWriter.Create(writer);
                entryXDoc.WriteTo(xmlWriter);
            }
        }
    }

}