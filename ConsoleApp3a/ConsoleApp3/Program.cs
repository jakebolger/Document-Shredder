using System;
using GroupDocs.Conversion;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;

namespace ConsoleApp3
{
    class Program
    {


        static void Main(string[] args)
        {
            // Load RTF file
            //
            var converter = new GroupDocs.Conversion.Converter(@"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\input_document.rtf");

            // Set conversion parameters for DOCX format
            //
            var convertOptions = converter.GetPossibleConversions()["docx"].ConvertOptions;

            // Convert to DOCX format
            //
            converter.Convert(@"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\output.docx", convertOptions);

            Document doc = new Document(@"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\output.docx");

            //getting the entire section
            //
            Section section = doc.Sections[0];

            //getting the first paragraph
            //
            Paragraph paragraph = section.Paragraphs[3];
            //Console.WriteLine(paragraph.Text);
            //getting the second paragraph
            //
            Paragraph paragraph2 = section.Paragraphs[7];

            //Console.WriteLine(paragraph2.Text);
            //getting the third paragraph
            //
            Paragraph paragraph3 = section.Paragraphs[12];
            //Console.WriteLine(paragraph3.Text);
            //getting the fourth paragraph
            //
            Paragraph paragraph4 = section.Paragraphs[16];
            //Console.WriteLine(paragraph4.Text);


            //Testing
            //
            /*
            Console.WriteLine(paragraph.Text);
            section.Paragraphs.RemoveAt(1);
            section.Paragraphs.RemoveAt(2);
            doc.SaveToFile(@"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\RemoveParagraphs.docx");
            
            //Section section = doc.Sections[0];
         
            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if (paragraph.StyleName == "sc_directional_language")
                    {
                    
                        Console.WriteLine(paragraph.Text);
                    }
                }
            }
            */

            string xmlFilePath = @"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\section_template.xml";

            XDocument xdoc = XDocument.Load(xmlFilePath);

            XNamespace ns = "http://www.w3.org/2001/XMLSchema-instance";

            //getting number at start of paragraph
            //
            string str1 = paragraph.Text;
            string str2 = string.Empty;
            int val = 0;

            //Console.WriteLine($"String with number: {str1}");
            //
            for (int i = 0; i < str1.Length; i++)
            {
                if (Char.IsDigit(str1[i]))
                    str2 += str1[i];
            }
            if (str2.Length > 0)
                val = int.Parse(str2);
            //Console.WriteLine($"Extracted Number: {val}");
            //
            
            //get first number / title
            //
            int firstDigit = (int)(val / Math.Pow(10, (int)Math.Floor(Math.Log10(val))));

            //second / chapter
            //
            int secondDigit = (val / 100) % 10;

            //third / number
            //
            int thirdDigit = (val / 10) % 10 * 10;


            //Console.WriteLine(thirdDigit);
            //Console.WriteLine(firstDigit);
            //Console.WriteLine(val);

            var query = from c in xdoc.Elements("root").Elements("section")
                        select c;

            
            foreach(XElement sectio in query)
            {
                sectio.Attribute("id").Value = $"{firstDigit}-{secondDigit}-{thirdDigit}";
                sectio.Attribute("title").Value = $"{firstDigit}";
                sectio.Attribute("chapter").Value = $"{secondDigit}";
                sectio.Attribute("number").Value = $"{thirdDigit}";

            }

            str1 = str1.Remove(0, 18);
            //Console.WriteLine(str1);
            //Console.WriteLine(xdoc);


            //XElement root = xdoc.Root;
            //Console.WriteLine(root);

            var trgt = xdoc.Root.Descendants("section").FirstOrDefault();
            var prg = trgt.Descendants("paragraph").FirstOrDefault();
            

            if (prg != null)
            {
                prg.Value = str1;
            }

                      
            //saving to xml file template
            //
            xdoc.Save(xmlFilePath);


            //next Xml file 2
            //================================================================================
            //

            string xmlFilePath2 = @"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\section_template2.xml";

            XDocument xdoc2 = XDocument.Load(xmlFilePath2);

            //XNamespace ns = "http://www.w3.org/2001/XMLSchema-instance";

            //getting number at start of paragraph
            //
            string str3 = paragraph2.Text;
            string str4 = string.Empty;
            int val2 = 0;

            //Console.WriteLine($"String with number: {str1}");
            //
            for (int i = 0; i < str3.Length; i++)
            {
                if (Char.IsDigit(str3[i]))
                    str4 += str3[i];
            }
            if (str4.Length > 0)
                val2 = int.Parse(str4);
            //Console.WriteLine($"Extracted Number: {val}");
            //
            
            //get first number / title
            int firstDigit2 = (int)(val2 / Math.Pow(10, (int)Math.Floor(Math.Log10(val2))));
            int secondDigit20 = (val2 / 1000) % 10;
            //second / chapter
            //
            int secondDigit2 = (val2 / 100) % 10;

            //third / number
            //
            int thirdDigit2 = (val2 / 10) % 10 * 10;
            //Console.WriteLine(val);

            var query2 = from c2 in xdoc2.Elements("root").Elements("section")
                        select c2;

            foreach (XElement sectio2 in query2)
            {
                
                sectio2.Attribute("id").Value = $"{firstDigit2}{secondDigit20}-{secondDigit2}-{thirdDigit2}";
                sectio2.Attribute("title").Value = $"{firstDigit2}{secondDigit20}";
                sectio2.Attribute("chapter").Value = $"{secondDigit2}";
                sectio2.Attribute("number").Value = $"{thirdDigit2}";
            }

            //Console.WriteLine(xdoc);


            str3 = str3.Remove(0, 19);
            //Console.WriteLine(str3);

            //XElement root = xdoc.Root;
            //Console.WriteLine(root);

            var trgt2 = xdoc2.Root.Descendants("section").FirstOrDefault();
            var prg2 = trgt2.Descendants("paragraph").FirstOrDefault();


            if (prg2 != null)
            {
                prg2.Value = str3;
            }


            //saving to xml file template
            //
            xdoc2.Save(xmlFilePath2);


            //next Xml file 3
            //================================================================================
            //

            string xmlFilePath3 = @"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\section_template3.xml";

            XDocument xdoc3 = XDocument.Load(xmlFilePath3);

            //XNamespace ns = "http://www.w3.org/2001/XMLSchema-instance";

            //getting number at start of paragraph
            //
            string str5 = paragraph3.Text;
            string str6 = string.Empty;
            int val3 = 0;
            //Console.WriteLine($"String with number: {str1}");
            //
            for (int i = 0; i < str5.Length; i++)
            {
                if (Char.IsDigit(str5[i]))
                    str6 += str5[i];
            }
            if (str6.Length > 0)
                val3 = int.Parse(str6);

            //Console.WriteLine($"Extracted Number: {val}");
            //

            //get first number / title
            int firstDigit3 = (int)(val3 / Math.Pow(10, (int)Math.Floor(Math.Log10(val3))));
            int secondDigit30 = (val3 / 10000) % 10;
            //second / chapter
            //
            int secondDigit3 = (val3 / 1000) % 10;
            int secondDigit301 = (val3 / 100) % 10;

            //third / number
            //
            int thirdDigit3 = (val3 / 10) % 10 * 10;

            //Console.WriteLine(val);

            var query3 = from c3 in xdoc3.Elements("root").Elements("section")
                         select c3;

            foreach (XElement sectio3 in query3)
            {
                
                sectio3.Attribute("id").Value = $"{firstDigit3}{secondDigit30}-{secondDigit3}{secondDigit301}-{thirdDigit3}";
                sectio3.Attribute("title").Value = $"{firstDigit3}{secondDigit30}";
                sectio3.Attribute("chapter").Value = $"{secondDigit3}{secondDigit301}";
                sectio3.Attribute("number").Value = $"{thirdDigit3}";
            }

            //Console.WriteLine(xdoc);


            str5 = str5.Remove(0, 20);
            //Console.WriteLine(str5);

            //XElement root = xdoc.Root;
            //Console.WriteLine(root);

            var trgt3 = xdoc3.Root.Descendants("section").FirstOrDefault();
            var prg3 = trgt3.Descendants("paragraph").FirstOrDefault();


            if (prg3 != null)
            {
                prg3.Value = str5;
            }


            //saving to xml file template
            //
            xdoc3.Save(xmlFilePath3);



            //next Xml file 4
            //================================================================================
            //

            string xmlFilePath4 = @"C:\Users\EKAJ\source\repos\ConsoleApp3\ConsoleApp3\section_template4.xml";

            XDocument xdoc4 = XDocument.Load(xmlFilePath4);

            //XNamespace ns = "http://www.w3.org/2001/XMLSchema-instance";

            //getting number at start of paragraph
            //
            string str7 = paragraph4.Text;
            string str8 = string.Empty;
            int val4 = 0;

            //Console.WriteLine($"String with number: {str1}");
            //
            for (int i = 0; i < str7.Length; i++)
            {
                if (Char.IsDigit(str7[i]))
                    str8 += str7[i];
            }
            if (str8.Length > 0)
                val4 = int.Parse(str8);

            //Console.WriteLine($"Extracted Number: {val}");
            //

            //get first number / title
            int firstDigit4 = (int)(val4 / Math.Pow(10, (int)Math.Floor(Math.Log10(val4))));
            int secondDigit40 = (val4 / 10000) % 10;
            //second / chapter
            //
            int secondDigit4 = (val4 / 1000) % 10;
            int secondDigit401 = (val4 / 100) % 10;

            //third / number
            //
            int thirdDigit4 = (val4 / 10) % 10 * 10;

            //Console.WriteLine(val);

            var query4 = from c4 in xdoc4.Elements("root").Elements("section")
                         select c4;

            foreach (XElement sectio4 in query4)
            {
                sectio4.Attribute("id").Value = $"{firstDigit4}{secondDigit40}-{secondDigit4}{secondDigit401}-{thirdDigit4}";
                sectio4.Attribute("title").Value = $"{firstDigit4}{secondDigit40}";
                sectio4.Attribute("chapter").Value = $"{secondDigit4}{secondDigit401}";
                sectio4.Attribute("number").Value = $"{thirdDigit4}";
            }

            //Console.WriteLine(xdoc);

            str7 = str7.Remove(0, 20);
            //Console.WriteLine(str7);


            //XElement root = xdoc.Root;
            //Console.WriteLine(root);

            var trgt4 = xdoc4.Root.Descendants("section").FirstOrDefault();
            var prg4 = trgt4.Descendants("paragraph").FirstOrDefault();


            if (prg4 != null)
            {
                prg4.Value = str7;
            }


            //saving to xml file template
            //
            xdoc4.Save(xmlFilePath4);

            Console.WriteLine("DOCUMENT SHREDDED. \n------------------\nXML FILES READY.");

        }
    }
}
