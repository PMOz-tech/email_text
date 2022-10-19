using System;
using System.Text;
using System.Xml;
using Aspose.Cells;

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


            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // The node is an element.
                        if (reader.Name != "see" && reader.Name != "c" && reader.Name != "paramref")
                        {
                            var title = (" " + reader.Name + "=>");
                            sb.Append(title);


                            while (reader.MoveToNextAttribute()) // Read the attributes.
                                sb.Append(reader.Name + "='" + reader.Value + "'");

                        }
                        //Console.Write(">");
                        //Console.WriteLine(">");
                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        sb.Append(reader.Value.Trim());
                        break;
                        // case XmlNodeType. EndElement: //Display the end of the element.
                        //     Console.Write("</" + reader.Name);
                        //     Console.Write("> \n");
                        //     break;
                }
            }
            // read and write to a file
            string folder = @"C:\Application Folder\email_text\bin\Debug\net6.0\";
            // Filename  
            string fileName = "TextFile.txt";

            string mdFileName = "TextResultDoc.md";
            string fullMdPath = folder + mdFileName;
           
            string fullPath = folder + fileName;

            File.WriteAllText(fullPath, sb.ToString());
            // Read a file  
            string readText = File.ReadAllText(fullPath);
            Console.WriteLine(readText);

            //using aspose nuget manager
            //var workbook = new Workbook(fullPath);
            //workbook.Save(fullMdPath);

            // using groupdocs convert
            var converter = new GroupDocs.Conversion.Converter(fullPath);
            // Prepare conversion options for target format MD
            var convertOptions = converter.GetPossibleConversions()["md"].ConvertOptions;
            // Convert to MD format
            converter.Convert(fullMdPath, convertOptions);
            Console.WriteLine(sb.ToString());
        }
    }
}
