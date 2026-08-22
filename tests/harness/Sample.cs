using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace DocumentTemplateXRay.Harness
{
    /// <summary>A template as the environment holds it, or as it sits on somebody's disk.</summary>
    public class SampleTemplate
    {
        public string Name;
        public string EntityType;

        /// <summary>The .docx this template is. A real one - the fixtures, or one built here.</summary>
        public string Path;
    }

    /// <summary>One table, as much of it as the display-name resolver ever looks at.</summary>
    public class SampleEntity
    {
        public string LogicalName;
        public string DisplayName;

        /// <summary>Logical name to display name.</summary>
        public readonly Dictionary<string, string> Attributes =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Relationship schema name to the table on the other end of it.</summary>
        public readonly Dictionary<string, string> ManyToOne =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public readonly Dictionary<string, string> OneToMany =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The environment the harness stands in for, and the templates in it.
    ///
    /// The templates are the fixtures in <c>fixtures\</c>, copied beside the harness at build
    /// time and served as real bytes: the tool unzips and parses them with the extractor it
    /// ships, so what a scenario reads off the screen is what the tool would really have found
    /// in that file.
    ///
    /// The metadata covers exactly the tables, columns and relationships those fixtures name, so
    /// a resolved field reads the way it would against a real Dataverse - and
    /// <c>04-unresolvable-fields.docx</c> stays unresolvable here too, because it names things no
    /// environment has.
    /// </summary>
    public static class Sample
    {
        /// <summary>Where the fixtures land beside the harness.</summary>
        public static string FixturesDir
        {
            get { return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fixtures"); }
        }

        public static string Fixture(string name)
        {
            return System.IO.Path.Combine(FixturesDir, name);
        }

        /// <summary>The templates this environment holds, in the order it hands them over.</summary>
        public static List<SampleTemplate> Templates()
        {
            return new List<SampleTemplate>
            {
                Template("Account Summary", "account", "01-duplicate-column-names.docx"),
                Template("Account with Contacts and Tasks", "account", "02-repeating-sections.docx"),
                Template("Account Letterhead", "account", "03-header-footer.docx"),
                Template("Account Audit (stale)", "account", "04-unresolvable-fields.docx"),
                Template("Blank Letter", "account", "05-no-fields.docx"),
            };
        }

        private static SampleTemplate Template(string name, string entity, string fixture)
        {
            return new SampleTemplate { Name = name, EntityType = entity, Path = Fixture(fixture) };
        }

        /// <summary>
        /// A template that is big rather than complicated: a handful of fields in a document
        /// part padded out to several megabytes, which is what a real letterhead with artwork and
        /// boilerplate looks like to the extractor.
        ///
        /// Big and simple on purpose. It takes a noticeable time to unzip and scan, and almost no
        /// time to draw - so what a scenario measures against it is the reading, not the drawing.
        /// </summary>
        public static string BigTemplate(string dir = null, int megabytes = 24)
        {
            dir = dir ?? System.IO.Path.Combine(System.IO.Path.GetTempPath(), "dtx-slow-harness");
            var path = System.IO.Path.Combine(dir, "big-letterhead.docx");
            if (File.Exists(path)) return path;

            Directory.CreateDirectory(dir);

            var document = new StringBuilder();
            document.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            document.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"");
            document.Append(" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w:body>");

            foreach (var field in new[] { "name", "description", "address1_city", "emailaddress1" })
            {
                document.Append(Sdt("account/" + field, field));
            }

            // The padding is text, not fields: the file is large, the answer is small. It has to
            // be incompressible text, or the zip would be a few kilobytes and nothing would
            // actually travel, be written or be read - which is the whole of what is measured.
            var random = new Random(20260821);
            var bytes = new byte[3000];
            for (var i = 0; document.Length < megabytes * 1024 * 1024; i++)
            {
                random.NextBytes(bytes);
                document.Append("<w:p><w:r><w:t>").Append(Convert.ToBase64String(bytes))
                    .Append("</w:t></w:r></w:p>");
            }

            document.Append("</w:body></w:document>");

            using (var file = File.Create(path))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                Write(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\" />"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd."
                    + "openxmlformats-officedocument.wordprocessingml.document.main+xml\" />"
                    + "</Types>");
                Write(zip, "word/document.xml", document.ToString());
            }

            return path;
        }

        /// <summary>One content control, bound the way a Dynamics template binds one.</summary>
        private static string Sdt(string path, string tag)
        {
            return "<w:sdt><w:sdtPr><w:alias w:val=\"" + tag + "\" /><w:tag w:val=\"\" />"
                   + "<w:dataBinding w:xpath=\"/" + path + "[1]\""
                   + " w:storeItemID=\"{A1B2C3D4-0000-0000-0000-000000000001}\" /></w:sdtPr>"
                   + "<w:sdtContent><w:r><w:t>" + tag + "</w:t></w:r></w:sdtContent></w:sdt>";
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using (var stream = zip.CreateEntry(name).Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        /// <summary>
        /// The tables the fixtures name, with the columns and relationship schema names they walk
        /// through. Anything not here resolves to nothing, which is what fixture 04 is for.
        /// </summary>
        public static Dictionary<string, SampleEntity> Metadata()
        {
            var account = Entity("account", "Account",
                "name:Account Name", "description:Description", "address1_city:City",
                "emailaddress1:Email", "telephone1:Main Phone", "createdon:Created On");
            account.ManyToOne["account_parent_account"] = "account";
            account.ManyToOne["account_primary_contact"] = "contact";
            account.ManyToOne["user_accounts"] = "systemuser";
            account.OneToMany["contact_customer_accounts"] = "contact";
            account.OneToMany["Account_Tasks"] = "task";

            var contact = Entity("contact", "Contact",
                "fullname:Full Name", "description:Description", "address1_city:City",
                "emailaddress1:Email", "telephone1:Business Phone", "jobtitle:Job Title",
                "createdon:Created On");

            var user = Entity("systemuser", "User",
                "fullname:Full Name", "title:Title", "address1_city:City");

            var task = Entity("task", "Task",
                "subject:Subject", "description:Description", "scheduledend:Due Date",
                "createdon:Created On");

            return new[] { account, contact, user, task }
                .ToDictionary(e => e.LogicalName, e => e, StringComparer.OrdinalIgnoreCase);
        }

        private static SampleEntity Entity(string logicalName, string displayName, params string[] attributes)
        {
            var entity = new SampleEntity { LogicalName = logicalName, DisplayName = displayName };
            foreach (var attribute in attributes)
            {
                var parts = attribute.Split(':');
                entity.Attributes[parts[0]] = parts[1];
            }

            return entity;
        }
    }
}
