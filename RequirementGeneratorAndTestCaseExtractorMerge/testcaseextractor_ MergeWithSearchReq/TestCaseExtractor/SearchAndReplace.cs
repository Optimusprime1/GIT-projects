using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    class SearchAndReplacer
    {
        public static XmlDocument GetXmlDocument(OpenXmlPart part)
        {
            XmlDocument xmlDoc = new XmlDocument();
            using (Stream partStream = part.GetStream())
            using (XmlReader partXmlReader = XmlReader.Create(partStream))
                xmlDoc.Load(partXmlReader);
            return xmlDoc;
        }

        public static void PutXmlDocument(OpenXmlPart part, XmlDocument xmlDoc)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                xmlDoc.Save(partXmlWriter);
        }


        static XmlNode SearchAndReplaceInWt(XmlNode rowlist ,string search)
        {
            XmlDocument xmlDoc = rowlist.OwnerDocument;
            string wordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", wordNamespace);
            XmlNodeList tlist = rowlist.SelectNodes("w:t", nsmgr);


            StringBuilder sb = new StringBuilder();
         
            
                foreach (XmlNode text in tlist)
                {
                    sb.Append(((XmlElement)text).InnerText);
                    if (sb.ToString().Contains(search))
                    {
                        //Console.WriteLine(sb);
                        return rowlist;
                    }
                   
                }

            rowlist = null;

            return rowlist;
        }

          
            /*
            if (sb.ToString().Contains(search) ||
                (!matchCase && sb.ToString().ToUpper().Contains(search.ToUpper())))

                                    if (sb.ToString().Contains("SRS"))
         
             */






        static void SearchAndReplaceInRow(XmlNode rowlist)
        {
            XmlDocument xmlDoc = rowlist.OwnerDocument;
            string wordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", wordNamespace);
            var Paraglists = rowlist.SelectNodes("descendant::w:p", nsmgr);


            XmlNode[] array = (new List<XmlNode>(Shim<XmlNode>(Paraglists))).ToArray();
          //  SearchAndReplaceInWt((XmlNode)array[1]);


            /*

            foreach (var paraglist in Paraglists)
            {
                SearchAndReplaceInWt((XmlElement)paraglist);
            }

            */

            //  XmlNodeList fistelement = paragraph.SelectNodes("attribute::*", nsmgr);

        }


        public static IEnumerable<T> Shim<T>(System.Collections.IEnumerable enumerable)
        {
            foreach (object current in enumerable)
            {
                yield return (T)current;
            }
        }

        public static string SearchAncestor(XmlNode row)
        {

            string myCollection = null;

            if (row != null)
            {
                XmlDocument xmlDoc = row.OwnerDocument;
                string wordNamespace =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("w", wordNamespace);
                XmlNodeList anch = row.SelectNodes("ancestor-or-self::w:tbl", nsmgr);
                StringBuilder sb = new StringBuilder();
                

                foreach (XmlNode text in anch)
                {

                    sb.Append(((XmlNode)text).InnerText);

                    // Console.WriteLine(sb.ToString());

                    string UnparsedString = sb.ToString();

                    int indexofTitle = UnparsedString.IndexOf("Title") + 5;



                    if (indexofTitle == 4)
                    {
                        int indexofUnique = UnparsedString.IndexOf("Unique");
                        int indexofoncelik = UnparsedString.IndexOf("Öncelik");
                        int indexofisim = UnparsedString.IndexOf("İsim") + 4;
                        int indexofSRS = UnparsedString.IndexOf("SRS.");

                        //Console.WriteLine(UnparsedString.Substring(indexofSRS, (indexofoncelik - indexofSRS)));

                        //Console.WriteLine(UnparsedString.Substring(indexofisim, (indexofUnique - indexofisim)));

                        //Fie io For testing isseus
                        /*
                        System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\Users\\vsanoglu\\Desktop\\test1.txt", true);
                        file.WriteLine(UnparsedString.Substring(indexofSRS, (indexofoncelik - indexofSRS)) + "-" + UnparsedString.Substring(indexofisim, (indexofUnique - indexofisim)));
                        file.Close();
                        
                        */
                        myCollection=UnparsedString.Substring(indexofSRS, (indexofoncelik - indexofSRS)) + "-" + UnparsedString.Substring(indexofisim, (indexofUnique - indexofisim));

                        
               



                    }
                    else
                    {

                        int indexofUnique = UnparsedString.IndexOf("Unique");
                        int indexofPriorty = UnparsedString.IndexOf("Priority");
                        int indexofSRS = UnparsedString.IndexOf("SRS.");
                        //Console.WriteLine(indexofPriorty.ToString(), "indexofPriorty");
                        //Console.WriteLine(indexofSRS.ToString(), "indexofSRS");
                       //Console.WriteLine(UnparsedString.Substring(indexofSRS, (indexofPriorty - indexofSRS)));

                        //Console.WriteLine(UnparsedString.Substring(indexofTitle, (indexofUnique - indexofTitle)));

                        // Write to file for testing purposes
                        /* 
                        System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\Users\\vsanoglu\\Desktop\\test1.txt", true);
                        file.WriteLine(UnparsedString.Substring(indexofSRS, (indexofPriorty - indexofSRS)) + "-" + UnparsedString.Substring(indexofTitle, (indexofUnique - indexofTitle)));

                        file.Close();

                        */

                        myCollection=UnparsedString.Substring(indexofSRS, (indexofPriorty - indexofSRS)) + "-" + UnparsedString.Substring(indexofTitle, (indexofUnique - indexofTitle));

                    }




                }


            }

            return myCollection;

        }
        static List<string> SearchAndReplaceInTr(XmlElement table, string search, bool matchCase)
        {
            XmlDocument xmlDoc = table.OwnerDocument;
            string wordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", wordNamespace);
            var rows = table.SelectNodes("descendant::w:r", nsmgr);
            List<string> myCollection = new List<string>();

            //XmlNode[] array = (new List<XmlNode>(Shim<XmlNode>(rows))).ToArray();
            // SearchAndReplaceInRow((XmlNode)array[2]);
            foreach (var row in rows)
            {
                
            string checknulll = SearchAncestor(SearchAndReplaceInWt((XmlElement)row, search));
                if (checknulll != null) { myCollection.Add(checknulll); }
               
            }

            /*
            foreach (var item in myCollection) {
                Console.WriteLine(item);
            }

            */
            return myCollection;

            /*

            foreach (var row in rows[1])
            {
                SearchAndReplaceInRow((XmlElement)row);
            }
         */




            //  XmlNodeList fistelement = paragraph.SelectNodes("attribute::*", nsmgr);



            //{
            //    XmlNodeList runs = paragraph.SelectNodes("child::w:r", nsmgr);
            //    foreach (XmlElement run in runs)
            //    {
            //        XmlNodeList childElements = run.SelectNodes("child::*", nsmgr);
            //        if (childElements.Count > 0)
            //        {
            //            XmlElement last = (XmlElement)childElements[childElements.Count - 1];
            //            for (int c = childElements.Count - 1; c >= 0; --c)
            //            {
            //                if (childElements[c].Name == "w:rPr")
            //                    continue;
            //                if (childElements[c].Name == "w:t")
            //                {
            //                    string textElementString = childElements[c].InnerText;
            //                    for (int i = textElementString.Length - 1; i >= 0; --i)
            //                    {
            //                        XmlElement newRun =
            //                            xmlDoc.CreateElement("w:r", wordNamespace);
            //                        XmlElement runProps =
            //                            (XmlElement)run.SelectSingleNode("child::w:rPr", nsmgr);
            //                        if (runProps != null)
            //                        {
            //                            XmlElement newRunProps =
            //                                (XmlElement)runProps.CloneNode(true);
            //                            newRun.AppendChild(newRunProps);
            //                        }
            //                        XmlElement newTextElement =
            //                            xmlDoc.CreateElement("w:t", wordNamespace);
            //                        XmlText newText =
            //                            xmlDoc.CreateTextNode(textElementString[i].ToString());
            //                        newTextElement.AppendChild(newText);
            //                        if (textElementString[i] == ' ')
            //                        {
            //                            XmlAttribute xmlSpace = xmlDoc.CreateAttribute(
            //                                "xml", "space",
            //                                "http://www.w3.org/XML/1998/namespace");
            //                            xmlSpace.Value = "preserve";
            //                            newTextElement.Attributes.Append(xmlSpace);
            //                        }
            //                        newRun.AppendChild(newTextElement);
            //                        paragraph.InsertAfter(newRun, run);
            //                    }
            //                }
            //                else
            //                {
            //                    XmlElement newRun = xmlDoc.CreateElement("w:r", wordNamespace);
            //                    XmlElement runProps =
            //                        (XmlElement)run.SelectSingleNode("child::w:rPr", nsmgr);
            //                    if (runProps != null)
            //                    {
            //                        XmlElement newRunProps =
            //                            (XmlElement)runProps.CloneNode(true);
            //                        newRun.AppendChild(newRunProps);
            //                    }
            //                    XmlElement newChildElement =
            //                        (XmlElement)childElements[c].CloneNode(true);
            //                    newRun.AppendChild(newChildElement);
            //                    paragraph.InsertAfter(newRun, run);
            //                }
            //            }
            //            paragraph.RemoveChild(run);
            //        }
            //    }

            //    while (true)
            //    {
            //        bool cont = false;
            //        runs = paragraph.SelectNodes("child::w:r", nsmgr);
            //        for (int i = 0; i <= runs.Count - search.Length; ++i)
            //        {
            //            bool match = true;
            //            for (int c = 0; c < search.Length; ++c)
            //            {
            //                XmlElement textElement =
            //                    (XmlElement)runs[i + c].SelectSingleNode("child::w:t", nsmgr);
            //                if (textElement == null)
            //                {
            //                    match = false;
            //                    break;
            //                }
            //                if (textElement.InnerText == search[c].ToString())
            //                    continue;
            //                if (!matchCase &&
            //                    textElement.InnerText.ToUpper() == search[c].ToString().ToUpper())
            //                    continue;
            //                match = false;
            //                break;
            //            }
            //            if (match)
            //            {
            //                XmlElement runProps =
            //                    (XmlElement)runs[i].SelectSingleNode("descendant::w:rPr", nsmgr);
            //                XmlElement newRun = xmlDoc.CreateElement("w:r", wordNamespace);
            //                if (runProps != null)
            //                {
            //                    XmlElement newRunProps = (XmlElement)runProps.CloneNode(true);
            //                    newRun.AppendChild(newRunProps);
            //                }
            //                XmlElement newTextElement =
            //                    xmlDoc.CreateElement("w:t", wordNamespace);
            //                XmlText newText = xmlDoc.CreateTextNode(replace);
            //                newTextElement.AppendChild(newText);
            //                if (replace[0] == ' ' || replace[replace.Length - 1] == ' ')
            //                {
            //                    XmlAttribute xmlSpace = xmlDoc.CreateAttribute("xml", "space",
            //                        "http://www.w3.org/XML/1998/namespace");
            //                    xmlSpace.Value = "preserve";
            //                    newTextElement.Attributes.Append(xmlSpace);
            //                }
            //                newRun.AppendChild(newTextElement);
            //                paragraph.InsertAfter(newRun, (XmlNode)runs[i]);
            //                for (int c = 0; c < search.Length; ++c)
            //                    paragraph.RemoveChild(runs[i + c]);
            //                cont = true;
            //                break;
            //            }
            //        }
            //        if (!cont)
            //            break;
            //    }

            //    // Consolidate adjacent runs that have only text elements, and have the
            //    // same run properties. This isn't necessary to create a valid document,
            //    // however, having the split runs is a bit messy.
            //    XmlNodeList children = paragraph.SelectNodes("child::*", nsmgr);
            //    List<int> matchId = new List<int>();
            //    int id = 0;
            //    for (int c = 0; c < children.Count; ++c)
            //    {
            //        if (c == 0)
            //        {
            //            matchId.Add(id);
            //            continue;
            //        }
            //        if (children[c].Name == "w:r" &&
            //            children[c - 1].Name == "w:r" &&
            //            children[c].SelectSingleNode("w:t", nsmgr) != null &&
            //            children[c - 1].SelectSingleNode("w:t", nsmgr) != null)
            //        {
            //            XmlElement runProps =
            //                (XmlElement)children[c].SelectSingleNode("w:rPr", nsmgr);
            //            XmlElement lastRunProps =
            //                (XmlElement)children[c - 1].SelectSingleNode("w:rPr", nsmgr);
            //            if ((runProps == null && lastRunProps != null) ||
            //                (runProps != null && lastRunProps == null))
            //            {
            //                matchId.Add(++id);
            //                continue;
            //            }
            //            if (runProps != null && runProps.InnerXml != lastRunProps.InnerXml)
            //            {
            //                matchId.Add(++id);
            //                continue;
            //            }
            //            matchId.Add(id);
            //            continue;
            //        }
            //        matchId.Add(++id);
            //    }

            //    for (int i = 0; i <= id; ++i)
            //    {
            //        var x1 = matchId.IndexOf(i);
            //        var x2 = matchId.LastIndexOf(i);
            //        if (x1 == x2)
            //            continue;
            //        StringBuilder sb2 = new StringBuilder();
            //        for (int z = x1; z <= x2; ++z)
            //            sb2.Append(((XmlElement)children[z]
            //                .SelectSingleNode("w:t", nsmgr)).InnerText);
            //        XmlElement newRun = xmlDoc.CreateElement("w:r", wordNamespace);
            //        XmlElement runProps =
            //            (XmlElement)children[x1].SelectSingleNode("child::w:rPr", nsmgr);
            //        if (runProps != null)
            //        {
            //            XmlElement newRunProps = (XmlElement)runProps.CloneNode(true);
            //            newRun.AppendChild(newRunProps);
            //        }
            //        XmlElement newTextElement = xmlDoc.CreateElement("w:t", wordNamespace);
            //        XmlText newText = xmlDoc.CreateTextNode(sb2.ToString());
            //        newTextElement.AppendChild(newText);
            //        if (sb2[0] == ' ' || sb2[sb2.Length - 1] == ' ')
            //        {
            //            XmlAttribute xmlSpace = xmlDoc.CreateAttribute(
            //                "xml", "space", "http://www.w3.org/XML/1998/namespace");
            //            xmlSpace.Value = "preserve";
            //            newTextElement.Attributes.Append(xmlSpace);
            //        }
            //        newRun.AppendChild(newTextElement);
            //        paragraph.InsertAfter(newRun, children[x2]);
            //        for (int z = x1; z <= x2; ++z)
            //            paragraph.RemoveChild(children[z]);
            //    }

            //    var txbxParagraphs = paragraph.SelectNodes("descendant::w:p", nsmgr);
            //    foreach (XmlElement p in txbxParagraphs)
            //        SearchAndReplaceInParagraph((XmlElement)p, search, replace, matchCase);
            //}
        }

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            XmlDocument doc = GetXmlDocument(part);
            string wordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("w", wordNamespace);
            string xpathExpression =
                "descendant::w:cellDel|" +
                "descendant::w:cellIns|" +
                "descendant::w:cellMerge|" +
                "descendant::w:customXmlDelRangeEnd|" +
                "descendant::w:customXmlDelRangeStart|" +
                "descendant::w:customXmlInsRangeEnd|" +
                "descendant::w:customXmlInsRangeStart|" +
                "descendant::w:del|" +
                "descendant::w:delInstrText|" +
                "descendant::w:delText|" +
                "descendant::w:ins|" +
                "descendant::w:moveFrom|" +
                "descendant::w:moveFromRangeEnd|" +
                "descendant::w:moveFromRangeStart|" +
                "descendant::w:moveTo|" +
                "descendant::w:moveToRangeEnd|" +
                "descendant::w:moveToRangeStart|" +
                "descendant::w:moveTo|" +
                "descendant::w:numberingChange|" +
                "descendant::w:rPrChange|" +
                "descendant::w:pPrChange|" +
                "descendant::w:rPrChange|" +
                "descendant::w:sectPrChange|" +
                "descendant::w:tcPrChange|" +
                "descendant::w:tblGridChange|" +
                "descendant::w:tblPrChange|" +
                "descendant::w:tblPrExChange|" +
                "descendant::w:trPrChange";
            XmlNodeList descendants = doc.SelectNodes(xpathExpression, nsmgr);
            return descendants.Count > 0;
        }

        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            if (PartHasTrackedRevisions(doc.MainDocumentPart))
                return false;
            foreach (var part in doc.MainDocumentPart.HeaderParts)
                if (PartHasTrackedRevisions(part))
                    return false;
            foreach (var part in doc.MainDocumentPart.FooterParts)
                if (PartHasTrackedRevisions(part))
                    return false;
            if (doc.MainDocumentPart.EndnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.EndnotesPart))
                    return false;
            if (doc.MainDocumentPart.FootnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.FootnotesPart))
                    return false;
            return false;
        }

        public static List<string> SearchAndReplaceInXmlDocument(XmlDocument xmlDocument, string search, bool matchCase)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDocument.NameTable);
            nsmgr.AddNamespace("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var paragraphs = xmlDocument.SelectNodes("descendant::w:p", nsmgr);
            List<string> myCollection = new List<string>();
            //string[] myCollection ;
            

            foreach (XmlNode paragraph in paragraphs)
            {
                myCollection.AddRange(SearchAndReplaceInTr((XmlElement)paragraph, search, matchCase));
              
            }

            return myCollection;


    }




        //   if (sb.ToString().Contains(search))
        //     Console.WriteLine(sb);

        /*
        foreach (XmlNode paragraph in a)
        {

            sb.Append(((XmlElement)paragraph).InnerText);
            if (sb.ToString().Contains(search))
            Console.WriteLine(sb);
        }
        */

        /*
        foreach (var paragraph in paragraphs)
            SearchAndReplaceInTable((XmlElement)paragraph, search, matchCase); }
*/

        public static List<string> SearchAndReplace(WordprocessingDocument wordDoc, string search,bool matchCase)
        {

            List<string> myCollection = new List<string>();
            if (HasTrackedRevisions(wordDoc))
                throw new SearchAndReplaceException(
                    "Search and replace will not work with documents " +
                    "that contain revision tracking.");

            XmlDocument xmlDoc;
            xmlDoc = GetXmlDocument(wordDoc.MainDocumentPart.DocumentSettingsPart);
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            XmlNodeList trackedRevisions =
                xmlDoc.SelectNodes("descendant::w:trackRevisions", nsmgr);
            

            /*
            if (trackedRevisions.Count > 0)
                throw new SearchAndReplaceException(
                    "Revision tracking is turned on for document.");
     */

            xmlDoc = GetXmlDocument(wordDoc.MainDocumentPart);
            myCollection.AddRange(SearchAndReplaceInXmlDocument(xmlDoc, search, matchCase));
            PutXmlDocument(wordDoc.MainDocumentPart, xmlDoc);

            /*
            
            if (wordDoc.MainDocumentPart.FootnotesPart != null)
            {
                xmlDoc = GetXmlDocument(wordDoc.MainDocumentPart.FootnotesPart);
                SearchAndReplaceInXmlDocument(xmlDoc, search, replace, matchCase);
                PutXmlDocument(wordDoc.MainDocumentPart.FootnotesPart, xmlDoc);
            }


            foreach (var part in wordDoc.MainDocumentPart.HeaderParts)
           {
               xmlDoc = GetXmlDocument(part);
               SearchAndReplaceInXmlDocument(xmlDoc, search, replace, matchCase);
               PutXmlDocument(part, xmlDoc);
           }

             foreach (var part in wordDoc.MainDocumentPart.FooterParts)
            {
                xmlDoc = GetXmlDocument(part);
                SearchAndReplaceInXmlDocument(xmlDoc, search, replace, matchCase);
                PutXmlDocument(part, xmlDoc);
            }
           
        if (wordDoc.MainDocumentPart.EndnotesPart != null)
            {
                xmlDoc = GetXmlDocument(wordDoc.MainDocumentPart.EndnotesPart);
                SearchAndReplaceInXmlDocument(xmlDoc, search, replace, matchCase);
                PutXmlDocument(wordDoc.MainDocumentPart.EndnotesPart, xmlDoc);
            }
     */

           

            return myCollection;
        }
        
    }

    public class SearchAndReplaceException : Exception
    {
        public SearchAndReplaceException(string message) : base(message) { }
    }









    public static class Extensions
    {
        public static string ToStringAlignAttributes(this XContainer xContainer)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            settings.NewLineOnAttributes = true;
            StringBuilder sb = new StringBuilder();
            using (XmlWriter xmlWriter = XmlWriter.Create(sb, settings))
                xContainer.WriteTo(xmlWriter);
            return sb.ToString();
        }

        public static XDocument GetXDocument(this XmlDocument document)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                document.WriteTo(xmlWriter);
            XmlDeclaration decl =
                document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding,
                    decl.Standalone);
            return xDoc;
        }

        public static XElement GetXElement(this XmlNode node)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                node.WriteTo(xmlWriter);
            return xDoc.Root;
        }
    }
}