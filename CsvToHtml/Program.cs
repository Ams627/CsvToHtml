using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace CsvToHtml
{
    public static class Extensions
    {
        public static void RunSta(this Thread thread)
        {
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }
    }
    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                string text = "";
                new Thread(() =>
                {
                    text = Clipboard.GetText();
                }).RunSta();

                var lines = Regex.Split(text, @"[\r\n]+");
                var htmlBuilder = new StringBuilder();
                htmlBuilder.AppendLine(@"<table class=""t1"">");
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    htmlBuilder.Append("<tr>");
                    var isHeader = line[0] == '.';
                    var lineToProcess = isHeader ? line.Substring(1) : line;
                    var cellTexts = lineToProcess.Split('|').Select(x => x.Trim());
                    foreach (var cell in cellTexts)
                    {
                        if (isHeader)
                        {
                            htmlBuilder.Append($"<th>{cell}</th>");
                        }
                        else
                        {
                            htmlBuilder.Append($"<td>{cell}</td>");
                        }
                    }
                    htmlBuilder.AppendLine("</tr>");
                }
                htmlBuilder.AppendLine(@"</table>");
                var html = htmlBuilder.ToString();
                var wf = GetHtmlClipboardContent(htmlBuilder.ToString());
                new Thread(() =>
                {
                    var dataObject = new DataObject();
                    dataObject.SetData(DataFormats.Html, wf);
                    dataObject.SetData(DataFormats.Text, html);
                    dataObject.SetData(DataFormats.UnicodeText, html);
                    Clipboard.SetDataObject(dataObject, true);
                }).RunSta();
                Console.WriteLine(wf);
            }
            catch (Exception ex)
            {
                var fullname = System.Reflection.Assembly.GetEntryAssembly().Location;
                var progname = Path.GetFileNameWithoutExtension(fullname);
                Console.Error.WriteLine($"{progname} Error: {ex.Message}");
            }

        }


        /// <summary>
        /// 1. start of html starting with the opening angle bracket
        /// 2. end of html is the length of the buffer
        /// 3. start fragment is the position of the start fragment comment starting with the opening angle bracket
        /// 4. end fragment is the position of the end fragment comment starting with the opening angle bracket
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        private static string GetHtmlClipboardContent(string html)
        {
            const string StartHtmlTag = "StartHtml:";
            const string EndHtmlTag = "StartHtml:";
            const string StartFragTag = "StartFragment:";
            const string EndFragTag = "EndFragment:";

            var startHtmlPosition = 0;
            var endHtmlPosition = 0;
            var startFragPosition = 0;
            var endFragPosition = 0;

            var fragment = new StringBuilder();
            for (int i = 0; i < 2; i++)
            {
                fragment.Clear();
                fragment.AppendLine("Version 0.9");
                fragment.AppendLine($"{StartHtmlTag}{startHtmlPosition:D8}");
                fragment.AppendLine($"{EndHtmlTag}{endHtmlPosition:D8}");
                fragment.AppendLine($"{StartFragTag}{startFragPosition:D8}");
                fragment.AppendLine($"{EndFragTag}{endFragPosition:D8}");

                fragment.AppendLine("<html><head><style>.t1, .t1 td, .t1 th {border:solid 1px gray;border-collapse:collapse}</style></head><body>");
                fragment.AppendLine("<!--StartFragment -->");
                fragment.AppendLine(html);
                fragment.AppendLine("<!--EndFragment-->");
                fragment.AppendLine("</body>");
                fragment.AppendLine("</html>");

                var str = fragment.ToString();
                startHtmlPosition = str.IndexOf("<html>");
                endHtmlPosition = str.Length;
                startFragPosition = str.IndexOf("<!--StartFragment");
                endFragPosition = str.IndexOf("<!--EndFragment");
            }
            return fragment.ToString();
        }
    }
}
