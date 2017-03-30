using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using System.Threading;
using System.IO;


namespace OpenXmlPowerTools
{
    class GetFileName
    {

        private const int NumberOfRetries = 3;
        private const int DelayOnRetry = 1000;

        public List<string> SelectFile(string file)
        {

         List<string> myCollection = new List<string>();

            for (int i = 1; i <= NumberOfRetries; ++i) {
                try {


                    using (WordprocessingDocument doc = WordprocessingDocument.Open(file, true))
                    //  WordprocessingDocument.Open("C:\\Users\\vsanoglu\\Desktop\\Requirement-project\\Microsoft CSM Application Analysis Document v2.3.docx", true))

                    {
                        //SearchAndReplacer.SearchAndReplace(doc, "Unique Req. Ref. No", true);
                        myCollection.AddRange(SearchAndReplacer.SearchAndReplace(doc, "Unique Req. Ref. No", true));

                    }
                    // Console.ReadKey();
                    //WordprocessingDocument.Open("C:\\Users\\vsanoglu\\Desktop\\Projects\\Eureko\\Eureko Sigorta_CRM_SRS_v1.0.docx", true))
                    break; // When done we can break loop

                }

         


    catch (IOException e) {
                    // You may check error code to filter some exceptions, not every error
                    // can be recovered.
                    if (i == NumberOfRetries) // Last one, (re)throw exception and exit
                    System.Windows.MessageBox.Show(e.Message);
                    Thread.Sleep(DelayOnRetry);
                    

                }

            }
            return myCollection;
        }

    }
}
