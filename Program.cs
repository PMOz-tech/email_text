using System;
using System.Globalization;
using System.Text;
using System.Xml;
// using Aspose.Cells;

namespace email_test
{
    /// <summary>
    /// Summary description for Class1.
    /// </summary>
    class Class1
    {
        static void Main(string[] args)
        {
            String URLString = "data.xml";
            XmlTextReader reader = new XmlTextReader(URLString);

            StringBuilder sb = new StringBuilder();
            List<string> contents = new List<string>();

            CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;
            TextInfo textInfo = cultureInfo.TextInfo;


            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // The node is an element.
                        if (reader.Name != "see" && reader.Name != "c" && reader.Name != "paramref" && reader.Name != "exception")
                        {
                            var title = (" " + reader.Name + ":");
                            sb.AppendLine(title);
                            if(title.Contains("member"))
                            {
                                contents.Add("\n");
                            }

                            contents.Add("**"+textInfo.ToTitleCase(title.Trim())+"**");


                            while (reader.MoveToNextAttribute()) // Read the attributes.
                                sb.AppendLine(reader.Value + "\n");
                                if(!string.IsNullOrEmpty(reader.Value))
                                {
                                    contents.Add("**"+reader.Value+"**\n");
                                }

                        }
                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        sb.Append(reader.Value.Trim());
                         if (reader.Name != "see" && reader.Name != "c" && reader.Name != "paramref")
                        {
                            contents.Add(reader.Value.Trim());
                        }
                        break;
                }
            }
          
            // read and write to a file
            string folder = @"C:\Application Folder\email_text\bin\Debug\net6.0\";
            // Filename  
            string fileName = "TextFile.txt";

            string mdFileName = "TextResultDoc.md";
            string fullMdPath = folder + mdFileName;
           
            string fullPath = folder + fileName;

            File.WriteAllLines(fullPath, contents);

            File.WriteAllLines(mdFileName, contents);
            // Read a file  
            // string readText = File.ReadAllText(fullPath);
            // Console.WriteLine(readText);

            //using aspose nuget manager
            // var workbook = new Workbook(fullPath);
            // workbook.Save(fullMdPath);

            // using groupdocs convert
            var converter = new GroupDocs.Conversion.Converter(fullPath);
            // Prepare conversion options for target format MD
            var convertOptions = converter.GetPossibleConversions()["md"].ConvertOptions;
            // Convert to MD format
            converter.Convert(fullMdPath, convertOptions);
            Console.WriteLine("completed");
        }
    }
}
