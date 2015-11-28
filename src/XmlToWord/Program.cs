
namespace XmlToWord
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml.Linq;
    using System.Xml.XPath;

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 1 && args[0].Equals("encodings", StringComparison.OrdinalIgnoreCase))
            {
                var encodings = Encoding.GetEncodings();
                for (int i = 0; i < encodings.Length; i++)
                {
                    Console.WriteLine(encodings[i].CodePage + "\t" + encodings[i].Name);
                }

                Console.WriteLine();
                Environment.ExitCode = 2;
                return;
            }

            if (args.Length < 5)
            {
                Console.WriteLine("Usage: encodings");
                Console.WriteLine("  Lists all available encodings");
                Console.WriteLine();
                Console.WriteLine("Usage: <xml in> <xml encoding> <word template> <word out> <items xpath> "
                    + "<{" + string.Join("|", Enum.GetNames(typeof(ItemStyle))) + "}{:item xpath:[attr name]:[format]}>+");
                Console.WriteLine("  Reads the XML file and create a Word document.");
                Console.WriteLine();
                Console.WriteLine("  Item XPath exemples:");
                Console.WriteLine("    > Header2:./Fields/Field[Name='Title']:Value");
                Console.WriteLine("    Finds the <Field> element using the Name attribute.");
                Console.WriteLine("    Takes the value of the Value attribute");
                Console.WriteLine();
                Console.WriteLine("    > Header2:./Fields/Field[Name='DateChanged']::DateChanged: {0}");
                Console.WriteLine("    Finds the <Field> element using the Name attribute.");
                Console.WriteLine("    Takes the value of the element.");
                Console.WriteLine("    The value is then formated to include a label in front of the value.");
                Console.WriteLine();
                Console.WriteLine("About <word template>");
                Console.WriteLine("  A template is necessary in order to use titles.");
                Console.WriteLine("  You can find the default template in a path like this one:");
                Console.WriteLine(@"  C:\Program Files (x86)\Microsoft Office\Office15\1033\QuickStyles\Default.dotx");
                Console.WriteLine();
                Environment.ExitCode = 1;
                return;
            }

            var context = new Context();
            var result = VerifyFile(context, args[0])
                && VerifyXmlEncoding(context, args[1])
                && VerifyWordTemplate(context, args[2])
                && ReadXmlFile(context, args[0])
                && PrepareWordFile(context, args[3], args[2])
                && ParseXml(context, args[4], args.Skip(5).ToArray())
                && CommitWordFile(context, args[3]);

            if (result)
            {
                Console.WriteLine("Done.");
                Environment.ExitCode = 0;
            }
            else
            {
                Environment.ExitCode = context.ExitCode ?? 99;
            }
        }

        private static bool VerifyFile(Context context, string filePath)
        {
            if (!File.Exists(filePath))
            {
                context.AddError("File does not exist '" + filePath + "'");
                context.ExitCode = 100;
            }

            context.Title = Path.GetFileNameWithoutExtension(filePath);

            return CheckErrors(context);
        }

        private static bool VerifyWordTemplate(Context context, string filePath)
        {
            if (!File.Exists(filePath))
            {
                context.AddError("File does not exist '" + filePath + "'");
                context.ExitCode = 101;
            }

            return CheckErrors(context);
        }

        private static bool VerifyXmlEncoding(Context context, string encoding)
        {
            var encodings = Encoding.GetEncodings();
            switch (encoding.ToLowerInvariant())
            {
                case "utf8":
                case "utf-8":
                    context.XmlEncoding = Encoding.UTF8;
                    break;

                default:
                    var match = encodings.FirstOrDefault(x => x.Name.Equals(encoding, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        context.XmlEncoding = Encoding.GetEncoding(match.CodePage);
                    }
                    else
                    {
                        context.ExitCode = 102;
                        context.AddError("Invalid encoding '" + encoding + "'.");
                    }

                    break;
            }

            return CheckErrors(context);
        }

        private static bool ReadXmlFile(Context context, string filePath)
        {
            try
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var reader = new StreamReader(file, context.XmlEncoding);
                    var doc = XDocument.Load(reader);
                    context.Xml = doc;
                }
            }
            catch (Exception ex)
            {
                context.ExitCode = 103;
                context.AddError("Failed to read XML file", ex.Message);
            }

            return CheckErrors(context);
        }

        private static bool PrepareWordFile(Context context, string filePath, string templateFilePath)
        {
            context.WordStream = new MemoryStream();

            try
            {
                var templateBytes = File.ReadAllBytes(templateFilePath);
                context.WordStream.Write(templateBytes, 0, templateBytes.Length);
                context.WordStream.Seek(0L, SeekOrigin.Begin);

                context.Word = WordprocessingDocument.Open(context.WordStream, true);
                context.Word.ChangeDocumentType(WordprocessingDocumentType.Document);
                ////var mainPart = context.WordPart = context.Word.AddMainDocumentPart();
                var mainPart = context.WordPart = context.Word.MainDocumentPart;
                var settings = mainPart.DocumentSettingsPart;

                ////var templateRelationship = new AttachedTemplate { Id = "relationId1" };
                ////settings.Settings.Append(templateRelationship);
                var templateUri = new Uri(templateFilePath, UriKind.Absolute);
                mainPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", templateUri, "relationId1");


                context.WordPart.Document = new Document();
                var body = context.WordPart.Document.AppendChild(new Body());
                context.WordBody = body;
                var h1Properties = new ParagraphProperties();
                h1Properties.Append(new ParagraphStyleId() { Val = "Heading1", });
                var h1 = body.AppendChild(new Paragraph());
                h1.Append(h1Properties);
                h1.AppendChild(new Run(new Text(context.Title ?? "Title 1")));
                //body.AppendChild(h1);
                body.AppendChild(new Paragraph(new Run(new Text("Document generated on " + DateTime.Now.ToString("o") + ". "))));
            }
            catch (Exception ex)
            {
                context.AddError("Failed to create in-memory Word document.", ex.Message);
                context.ExitCode = 104;
            }

            return CheckErrors(context);
        }

        private static bool ParseXml(Context context, string rootPath, string[] itemsPath)
        {
            var collection = context.Xml.Root.XPathSelectElements(rootPath);
            if (collection == null)
            {
                context.AddError("Cound not find root element from path '" + rootPath + "'");
                context.ExitCode = 105;
                return false;
            }

            context.ExitCode = 106;
            List<ItemPath> paths;
            try
            {
                paths = ParseXmlArguments(context, itemsPath);
            }
            catch (Exception ex)
            {
                context.AddError("An error occured while parsing the arguments.", ex.Message);
                context.ExitCode = 108;
                return CheckErrors(context);
            }

            try
            {
                ReadXmlAndWriteWord(context, collection, paths);
            }
            catch (Exception ex)
            {
                context.AddError("An error occured while generating the document content.", ex.Message);
                context.ExitCode = 109;
            }

            return CheckErrors(context);
        }

        private static List<ItemPath> ParseXmlArguments(Context context, string[] itemsPath)
        {
            var paths = new List<ItemPath>(itemsPath.Length);
            var chars = new char[] { ':', ';', };
            foreach (var spec in itemsPath)
            {
                var path = new ItemPath();

                if (spec.Count(c => c.Equals(':')) >= 3)
                {
                    var parts = spec.Split(new char[] { ':', }, 4);

                    ItemStyle style;
                    if (Enum.TryParse<ItemStyle>(parts[0], out style))
                    {
                        path.Style = style;
                    }
                    else
                    {
                        context.AddError("Invalid style in '" + spec + "'.");
                    }

                    path.Path = parts[1];
                    path.Attribute = string.IsNullOrWhiteSpace(parts[2]) ? null : parts[2];
                    path.Text = string.IsNullOrEmpty(parts[3]) ? null : parts[3];
                }
                else
                {
                    context.AddError("Not enough parameters in '" + spec + "'.");
                    continue;
                }

                paths.Add(path);
            }
            return paths;
        }

        private static void ReadXmlAndWriteWord(Context context, IEnumerable<XElement> collection, List<ItemPath> paths)
        {
            foreach (var element in collection)
            {
                foreach (var path in paths)
                {
                    string value = string.Empty;
                    if (path.Path != null)
                    {
                        var valueElement = element.XPathSelectElement(path.Path);

                        if (path.Attribute != null)
                        {
                            value = valueElement.Attribute(path.Attribute).Value;
                        }
                        else
                        {
                            value = valueElement.Value;
                        }

                        if (path.Text != null)
                        {
                            value = string.Format(path.Text, value);
                        }
                    }
                    else if (path.Text != null)
                    {
                        value = path.Text;
                    }

                    var p = new Paragraph();
                    var props = p.ParagraphProperties = new ParagraphProperties();
                    props.Append(new ParagraphStyleId() { Val = path.Style.ToString(), });

                    var lines = value.Split('\n');
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (i > 0)
                        {
                            p.AppendChild(new Break());
                        }

                        p.AppendChild(new Run(new Text(lines[i])));
                    }

                    context.WordBody.AppendChild(p);
                }
            }
        }

        private static bool CommitWordFile(Context context, string filePath)
        {
            context.WordPart.Document.Save();
            context.Word.Close();

            try
            {
                var bytes = context.WordStream.ToArray();
                File.WriteAllBytes(filePath, bytes);
            }
            catch (Exception ex)
            {
                context.ExitCode = 107;
                context.AddError("Failed to save word document.", ex.Message);
            }

            return CheckErrors(context);
        }

        private static bool CheckErrors(Context context)
        {
            if (context.Errors.Count > 0)
            {
                Console.WriteLine("Errors occured");
                Console.WriteLine("===============");
                Console.WriteLine();
                for (int i = 0; i < context.Errors.Count; i++)
                {
                    Console.WriteLine(context.Errors[i].Message);
                    if (context.Errors[i].Detail != null)
                    {
                        Console.WriteLine(context.Errors[i].Detail);
                    }

                    Console.WriteLine();
                }

                Console.WriteLine();
                return false;
            }

            return true;
        }
    }
}
