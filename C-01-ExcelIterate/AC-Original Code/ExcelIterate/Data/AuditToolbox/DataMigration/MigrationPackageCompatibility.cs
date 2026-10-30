using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Excel_Iterate
{
    // Pure file processing, restricted by callers to their owned disposable
    // copies. Old WPS reads OOXML ISO dates (t="d") as just the year number.
    internal static class MigrationPackageCompatibility
    {
        private static XmlReaderSettings Settings()
        { return new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = true, MaxCharactersInDocument = 64L * 1024 * 1024 }; }
        private static bool IsZip(Stream stream)
        { bool zip = stream.ReadByte() == 'P' && stream.ReadByte() == 'K'; stream.Position = 0; return zip; }
        private static bool HasDates(ZipArchiveEntry entry)
        {
            using (var reader = XmlReader.Create(entry.Open(), Settings()))
                while (reader.Read()) if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c" && reader.GetAttribute("t") == "d") return true;
            return false;
        }
        internal static bool HasIsoDates(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (!IsZip(stream)) return false;
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
                    return zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).Any(HasDates);
            }
        }
        internal static double DateSerial(string text, bool date1904)
        {
            DateTime date;
            if (!Regex.IsMatch(text, @"^\d{4}-\d{2}-\d{2}(?:T.*)?$") || !DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date))
                throw new InvalidOperationException("无法可靠读取 ISO 日期：" + text);
            date = DateTime.SpecifyKind(date, DateTimeKind.Unspecified);
            if (date1904) return (date - new DateTime(1904, 1, 1)).TotalDays;
            return (date - new DateTime(1899, 12, 31)).TotalDays + (date >= new DateTime(1900, 3, 1) ? 1 : 0);
        }
        internal static int NormalizeOwnedCopy(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                if (!IsZip(stream)) return 0;
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Update))
                {
                    bool date1904 = false; var workbook = zip.GetEntry("xl/workbook.xml");
                    if (workbook == null) return 0;
                    using (var reader = XmlReader.Create(workbook.Open(), Settings()))
                        while (reader.Read()) if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "workbookPr")
                        { string value = reader.GetAttribute("date1904"); date1904 = value == "1" || value == "true"; }
                    int changed = 0;
                    foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).ToArray())
                    {
                        if (!HasDates(entry)) continue;
                        var document = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
                        using (var reader = XmlReader.Create(entry.Open(), Settings())) document.Load(reader);
                        var ns = new XmlNamespaceManager(document.NameTable); ns.AddNamespace("s", document.DocumentElement.NamespaceURI);
                        foreach (XmlElement cell in document.SelectNodes("//s:sheetData/s:row/s:c[@t='d']", ns))
                        {
                            var value = cell.SelectSingleNode("s:v", ns);
                            if (value == null || cell.SelectSingleNode("s:f", ns) != null) throw new InvalidOperationException("ISO 日期单元格结构不受支持：" + cell.GetAttribute("r"));
                            value.InnerText = DateSerial(value.InnerText, date1904).ToString("R", CultureInfo.InvariantCulture);
                            cell.SetAttribute("t", "n"); changed++;
                        }
                        using (var output = entry.Open())
                        {
                            output.SetLength(0);
                            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false })) document.Save(writer);
                        }
                    }
                    return changed;
                }
            }
        }
    }
}
