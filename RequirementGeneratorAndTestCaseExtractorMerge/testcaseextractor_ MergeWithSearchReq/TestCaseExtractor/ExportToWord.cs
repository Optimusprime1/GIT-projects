using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Novacode;
using System.Diagnostics;


namespace TestCaseExtractor
{
    public class ExportToWord : MainWindow
    {
     
  internal static void Access_word(ITestSuiteBase rootSuite, ITestManagementTeamProject _testProject, string TbFileNameForExcel)
        {

            string fileName = @TbFileNameForExcel + ".docx";

            // Create a document in memory:
            var doc = DocX.Create(fileName);

            doc.Save();

        }
        


    }
}
