using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace App.Api.Processor.ExtractorService.Word
{
    
    public class WordFixer
    {
        public static void Fix(string filePath)
        {
            try
            {
                using (WordprocessingDocument.Open(filePath, true)) {}
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("malformed URI"))
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Fix problematic URI's
                        FixInvalidUri(fs, brokenUri => FixUri());
                    }
                }
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true))
            {
                MainDocumentPart part = wordDocument.MainDocumentPart;
                string xml = ReadString(part);
                XElement document = XElement.Parse(xml);

                XElement transformedDocument = (XElement) StripSmartTags(document);
                while (transformedDocument.Descendants().Where(d => d.Name.LocalName == "smartTag").FirstOrDefault() != null)
                {
                    transformedDocument = (XElement) StripSmartTags(transformedDocument);
                }

                bool ttt = transformedDocument.Descendants().Where(d => d.Name.LocalName == "smartTag").FirstOrDefault() != null;
                // Write the transformed document back to the part.
                WriteString(part, transformedDocument.ToString(SaveOptions.DisableFormatting));
            }
        }

        private static void WriteString(OpenXmlPart part, string text)
        {
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var streamWriter = new StreamWriter(stream);
            streamWriter.Write(text);
        }

        private static string ReadString(OpenXmlPart part)
        {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var streamReader = new StreamReader(stream);
            return streamReader.ReadToEnd();
        }

        private static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                foreach (var entry in za.Entries.ToList())
                {
                    if (!entry.Name.EndsWith(".rels"))
                        continue;
                    bool replaceEntry = false;
                    XDocument entryXDoc = null;
                    using (var entryStream = entry.Open())
                    {
                        try
                        {
                            entryXDoc = XDocument.Load(entryStream);
                            if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                            {
                                var urisToCheck = entryXDoc
                                    .Descendants(relNs + "Relationship")
                                    .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                foreach (var rel in urisToCheck)
                                {
                                    var target = (string)rel.Attribute("Target");
                                    if (target != null)
                                    {
                                        try
                                        {
                                            Uri uri = new Uri(target);
                                        }
                                        catch (UriFormatException)
                                        {
                                            Uri newUri = invalidUriHandler(target);
                                            rel.Attribute("Target").Value = newUri.ToString();
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
                        var fullName = entry.FullName;
                        entry.Delete();
                        var newEntry = za.CreateEntry(fullName);
                        using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                        using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                        {
                            entryXDoc.WriteTo(xmlWriter);
                        }
                    }
                }
            }
        }

        private static Uri FixUri()
        {
            return new Uri("http://broken-link/");
        }

        /// <summary>
        /// Recursive, pure functional transform that removes all w:smartTag elements.
        /// </summary>
        /// <param name="node">The <see cref="XNode" /> to be transformed.</param>
        /// <returns>The transformed <see cref="XNode" />.</returns>
        private static object StripSmartTags(XNode node)
        {
            // We only consider elements (not text nodes, for example).
            if (!(node is XElement element))
            {
                return node;
            }

            // Strip w:smartTag elements by only returning their children.
            if (element.Name.LocalName == "smartTag")
            {
                return element.Elements();
            }

            // Perform the identity transform.
            return new XElement(element.Name, element.Attributes(),
                element.Nodes().Select(StripSmartTags));
        }
    }
}
