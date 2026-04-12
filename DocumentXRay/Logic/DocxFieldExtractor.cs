using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocumentXRay.Logic
{
    public static class DocxFieldExtractor
    {
        private static readonly Regex XmlPartPattern =
            new Regex(@"^word/(document|header\d*|footer\d*)\.xml$", RegexOptions.Compiled);

        private static readonly Regex SdtPrPattern =
            new Regex(@"(?s)<w:sdtPr\b[^>]*>(.*?)</w:sdtPr>", RegexOptions.Compiled);

        private static readonly Regex XPathAttrPattern =
            new Regex(@"w:xpath=""([^""]+)""", RegexOptions.Compiled);

        private static readonly Regex StoreIdPattern =
            new Regex(@"w:storeItemID=""([^""]+)""", RegexOptions.Compiled);

        private static readonly Regex TagPattern =
            new Regex(@"<w:tag\s+w:val=""([^""]*)""", RegexOptions.Compiled);

        private static readonly Regex AliasPattern =
            new Regex(@"<w:alias\s+w:val=""([^""]*)""", RegexOptions.Compiled);

        private static readonly Regex IndexPattern =
            new Regex(@"\[\d+\]", RegexOptions.Compiled);

        private static readonly Regex NsPattern =
            new Regex(@"ns\d+:", RegexOptions.Compiled);

        private static readonly Regex RepeatingSectionPattern =
            new Regex(@"<w15:repeatingSection\s*/>", RegexOptions.Compiled);

        public static List<FieldInfo> ExtractFields(string docxPath)
        {
            var xmlParts = ExtractXmlParts(docxPath);
            var allFields = new List<FieldInfo>();

            foreach (var kvp in xmlParts)
            {
                var fields = ParseContentControls(kvp.Value, kvp.Key);
                allFields.AddRange(fields);
            }

            // Collect repeating section paths, then tag child fields
            var repeatingSectionPaths = allFields
                .Where(f => f.IsRepeatingSection && f.FieldPath != null)
                .Select(f => f.FieldPath)
                .ToList();

            foreach (var field in allFields)
            {
                if (field.IsRepeatingSection || field.FieldPath == null) continue;

                foreach (var rsPath in repeatingSectionPaths)
                {
                    if (field.FieldPath.StartsWith(rsPath + "/"))
                    {
                        field.RepeatingSectionName = rsPath.Contains("/")
                            ? rsPath.Substring(rsPath.LastIndexOf('/') + 1)
                            : rsPath;
                        break;
                    }
                }
            }

            return allFields;
        }

        private static Dictionary<string, string> ExtractXmlParts(string docxPath)
        {
            var parts = new Dictionary<string, string>();

            using (var zip = ZipFile.OpenRead(docxPath))
            {
                foreach (var entry in zip.Entries)
                {
                    if (!XmlPartPattern.IsMatch(entry.FullName))
                        continue;

                    using (var stream = entry.Open())
                    using (var reader = new StreamReader(stream))
                    {
                        parts[entry.FullName] = reader.ReadToEnd();
                    }
                }
            }

            return parts;
        }

        private static string ConvertXPathToFieldPath(string xpath)
        {
            var cleaned = IndexPattern.Replace(xpath, "");
            cleaned = NsPattern.Replace(cleaned, "");

            var segments = cleaned.TrimStart('/').Split('/')
                .Where(s => s != "DocumentTemplate" && s.Length > 0)
                .ToArray();

            return segments.Length == 0 ? null : string.Join("/", segments);
        }

        private static List<FieldInfo> ParseContentControls(string xmlContent, string partName)
        {
            var fields = new List<FieldInfo>();

            foreach (Match sdtPrMatch in SdtPrPattern.Matches(xmlContent))
            {
                var prContent = sdtPrMatch.Groups[1].Value;

                var xpathMatch = XPathAttrPattern.Match(prContent);
                if (!xpathMatch.Success) continue;

                var xpathValue = xpathMatch.Groups[1].Value;
                var storeIdMatch = StoreIdPattern.Match(prContent);
                var tagMatch = TagPattern.Match(prContent);
                var aliasMatch = AliasPattern.Match(prContent);

                var isRepeating = RepeatingSectionPattern.IsMatch(prContent);

                fields.Add(new FieldInfo
                {
                    FieldPath = ConvertXPathToFieldPath(xpathValue),
                    Tag = tagMatch.Success ? tagMatch.Groups[1].Value : null,
                    Alias = aliasMatch.Success ? aliasMatch.Groups[1].Value : null,
                    XPath = xpathValue,
                    StoreId = storeIdMatch.Success ? storeIdMatch.Groups[1].Value : null,
                    Location = partName,
                    IsRepeatingSection = isRepeating
                });
            }

            return fields;
        }
    }
}
