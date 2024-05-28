using System.IO.Compression;
using System.Xml;

Console.Write("File: ");
var file = new FileInfo(Console.ReadLine()!);
Console.Write("Attachments: ");
var attachments = new DirectoryInfo(Console.ReadLine()!);
Console.Write("Output: ");
var output = file.CopyTo(Console.ReadLine()!, false);

using var zip = ZipFile.Open(output.FullName, ZipArchiveMode.Update);
var contentTypes = zip.GetEntry("[Content_Types].xml");

XmlDocument xml = new XmlDocument();
using (var stream = contentTypes!.Open())
    xml.Load(stream);

var document = xml.DocumentElement;
foreach (var attachment in attachments.EnumerateFiles("*", SearchOption.AllDirectories))
{
    var path = Path.GetRelativePath(attachments.FullName, attachment.FullName);
    path = Path.Join("oxa", path);
    path = path.Replace('\\', '/');

    var element = xml.CreateElement("Override", document!.NamespaceURI);
    element.SetAttribute("PartName", $"/{path}");
    element.SetAttribute("ContentType", "application/octet-stream");
    _ = document.AppendChild(element);

    _ = zip.CreateEntryFromFile(attachment.FullName, path, CompressionLevel.SmallestSize);
}

contentTypes.Delete();
contentTypes = zip.CreateEntry("[Content_Types].xml");
using (var stream = contentTypes!.Open())
    xml.Save(stream);
