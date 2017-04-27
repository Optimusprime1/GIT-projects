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


        internal static DocX Access_word(ITestSuiteBase rootSuite, ITestManagementTeamProject _testProject, string TbFileNameForWord)
        {

            string fileName = @TbFileNameForWord + ".docx";

            // Create a document in memory:
            var doc = DocX.Create(fileName);

            return doc;

        }


            /*
             internal static void  WriteTestCaseToWord(ITestBase testCase, ITestManagementTeamProject _testProject , TfsTeamProjectCollection _tfs, DocX doc)
                    {


                        string testcaseid = testCase.Id.ToString();



                        var testResults = _testProject.TestResults.ByTestId(testCase.Id);

                        WorkItemStore Store = _tfs.GetService<WorkItemStore>();


                        foreach (ITestCaseResult result in testResults)
                        {

                            string testcase_result = null;

                            if (result.Outcome.ToString() == "Passed")

                            {  }

                            else if (result.Outcome.ToString() == "Failed")

                            { worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.Red); }

                            else if (result.Outcome.ToString() == "NotApplicable") { worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.Gray); }
                            else if (result.Outcome.ToString() == "Blocked") { worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.LightPink); }
                            else
                            {
                                worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke); worksheet.Cells[_i, 5].Value = "Design";

                            }

                        }



                        WorkItemCollection BugQueryResults = Store.Query(
                        "SELECT [System.Id],[Microsoft.VSTS.Common.StackRank],[Microsoft.VSTS.Common.Priority], [System.WorkItemType],[System.Title] " +
                        "From WorkItems " +
                        "Where [System.TeamProject] = '" + _testProject.WitProject.Name + "'" + "AND ([System.WorkItemType] = 'Bug' OR [System.WorkItemType] = 'Requirement' ) AND  [System.IterationPath] = '" + testCase.WorkItem.IterationPath + "' AND[System.State] <> 'Closed'");




                        /*

                        foreach (FieldDefinition disp in BugQueryResults.DisplayFields)
                        {

                            Console.WriteLine(disp.Name);

                        }

                        */

            /*
                        WorkItem myitem = testCase.WorkItem;

                        string _lastvalue = null;

                        var j = 1;


                        /* For download code
                        var wi = Store.GetWorkItem(testCase.Id);
                        string name =null;
                        Attachment attach = null;
                        attach = wi.Attachments.Cast<Attachment>().FirstOrDefault(x => x.Name == name);

                        */

            /*
                        foreach (ITestAction action in testCase.Actions)
                        {
                            var sharedRef = action as ISharedStepReference;

                            string testAction;
                            string expectedResult;
                            if (sharedRef != null)
                            {
                                var sharedStep = sharedRef.FindSharedStep();
                                foreach (var testStep in sharedStep.Actions.Select(sharedAction => sharedAction as ITestStep))
                                {
                                    testAction = j.ToString() + ". " +
                                                 ((testStep.Title.ToString().Length == 0)
                                                     ? "<<Not Recorded>>"
                                                     : testStep.Title.ToString());
                                    expectedResult = ((testStep.ExpectedResult.ToString().Length == 0)
                                        ? "<<Not Recorded>>"
                                        : testStep.ExpectedResult.ToString());
                                    WriteTestStepsToExcel(worksheet, _i, j, StripTagsCharArray(testAction), StripTagsCharArray(expectedResult));
                                    j++;
                                }
                            }
                            else
                            {
                                var testStep = action as ITestStep;
                                testAction = j.ToString() + ". " +
                                             ((testStep.Title.ToString().Length == 0)
                                                 ? "<<Not Recorded>>"
                                                 : testStep.Title.ToString());
                                expectedResult = ((testStep.ExpectedResult.ToString().Length == 0)
                                    ? "<<Not Recorded>>"
                                    : testStep.ExpectedResult.ToString());

                                WriteTestStepsToExcel(worksheet, _i, j, StripTagsCharArray(testAction), StripTagsCharArray(expectedResult));
                                j++;
                            }
                        } //end of foreach test action
                        _i = _i + j - 1;




                    }



                */

        }


}
