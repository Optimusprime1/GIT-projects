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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenXmlPowerTools;
using System.Collections.ObjectModel;
using Novacode;
using System.Diagnostics;
using System.Threading.Tasks;

namespace TestCaseExtractor
{
    /// <summary>
    /// Created at 18.01.2017 By Volkan Sanoglu for Veripark otomatic Requirement Generation from uploaded Documement
    public class MylistElements 
    {
        public string RqName { get; set; }


        public MylistElements(string Rqname)
        {
            this.RqName = Rqname;

        }

    }

    public partial class MainWindow 
    {
        private TfsTeamProjectCollection _tfs;
        ITestManagementTeamProject _testProject;
        ITestPlanCollection _testPlanCollection;
        private ITestSuiteBase _suite;
        int _i = 3;
        public System.Windows.Controls.DataGrid sender;
        List<string> myCollection = new List<string>();
        public ObservableCollection<MylistElements> Datagriditems;
        public string  rootsuiteid;
        int  _tvItem;
        int LoadcomboBoxindex = -1;
        string absolutepath = null;


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadcomboBox.Items.Add("Export Test Cases");
            LoadcomboBox.Items.Add("Create Requirements From Analysis");
          //  LoadcomboBox.Items.Add("Generate Proof Document");
        }

        private void BtnConnectForExcel_Click(object sender, RoutedEventArgs e)
        {

            _tfs = null;
            TbTfs.Text = null;
            var tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            if (tpp.SelectedTeamProjectCollection == null) return;

            _tfs = tpp.SelectedTeamProjectCollection;

            var testService = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));

            _testProject = testService.GetTeamProject(tpp.SelectedProjects[0].Name);

            TbTfs.Text = _tfs.Name + "\\" + _testProject;

            _testPlanCollection = _testProject.TestPlans.Query("Select * from TestPlan");

            ProjectSelected_GetTestPlans();

        }


        public string GetTestPlanName(ITestPlan selectedTestPlan)
        {

            string RootSuiteName = null;
            string _suitename = null;

            if (TvSuites.SelectedValue == null)
            {
                System.Windows.MessageBox.Show("Please Select Test Suite First");
                return null;
            }

            else { 

            RootSuiteName = selectedTestPlan.RootSuite.Title;

            var tvItem = TvSuites.SelectedItem as TreeViewItem;

             _suitename = tvItem.Header.ToString();

            // _suite = _testProject.TestSuites.Find(Convert.ToInt32(tvItem.Tag.ToString()));

            
            }

            return RootSuiteName + "_" + _suitename + "_" + DateTime.Now.ToString("dd.MM.yyyy");

        }


        private void ProjectSelected_GetTestPlans()
        {
            LbSelectTestPlan.ItemsSource = _testPlanCollection;
            LbSelectTestPlan.DisplayMemberPath = NameProperty.ToString();
            LbSelectTestPlan.SelectedIndex = -1;
        }


        private void LbSelectTestPlan_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TvSuites.Items.Clear();
            GetTestSuites(LbSelectTestPlan.SelectedItem as ITestPlan,null);
            


         //   rootsuiteid = LbSelectTestPlan.SelectedItem.ToString();

        }


       public void GetTestSuites(ITestPlan selectedTestPlan , string allitemsHeader)
        {
            if (selectedTestPlan == null) return;
            var root = new TreeViewItem { Header = selectedTestPlan.RootSuite.Title };
            TvSuites.Items.Add(root);
            root.Tag = selectedTestPlan.RootSuite.Id;
            

            GetSubSuites(selectedTestPlan.RootSuite.SubSuites, root, allitemsHeader);
           
            _tvItem = selectedTestPlan.Id;

        }

        static void GetSubSuites(IEnumerable<ITestSuiteBase> subSuiteEntries, ItemsControl treeNode,string allitemsHeader)
        {
            foreach (var suite in subSuiteEntries)
            {
                var suiteTree = new TreeViewItem { Header = suite.Title };
                treeNode.Items.Add(suiteTree);
                suiteTree.Tag = suite.Id;
                if (suite.TestSuiteType == TestSuiteType.StaticTestSuite)
                {
                    var suite1 = suite as IStaticTestSuite;
                    if (suite1.SubSuites.Count > 0)
                    {
                        GetSubSuites(suite1.SubSuites, suiteTree, allitemsHeader);
                    }
                }
            }

            


        }


        static string FindRootSuite(ITestSuiteBase currentSuite, string suiteTreePath)
        {
            var parentSuite = currentSuite.Parent;
            if (parentSuite == null) return suiteTreePath;
            if (parentSuite.IsRoot) return suiteTreePath;
            suiteTreePath = currentSuite.Parent.Title + ">" + suiteTreePath;
            var path = FindRootSuite(parentSuite, suiteTreePath);
            return path;
        }

        static string FindRootSuite(IRequirementTestSuite currentSuite, string suiteTreePath)
        {
            var parentSuite = currentSuite.Parent;
            if (parentSuite == null) return suiteTreePath;
            if (parentSuite.IsRoot) return suiteTreePath;
            suiteTreePath = currentSuite.Parent.Title + ">" + suiteTreePath;
            var path = FindRootSuite(parentSuite, suiteTreePath);
            return path;
        }


        void Access_Documents(ITestSuiteBase rootSuite, ITestManagementTeamProject _testProject,string allitemsHeader)
        {
            if (LoadcomboBoxindex == 2)
            {
            DocX worddoc =ExportToWord.Access_word(rootSuite, _testProject, TbFileNameForExcel.Text);
                GetTestSuites(rootSuite as IStaticTestSuite, null,null,worddoc, _testProject, allitemsHeader);

                worddoc.InsertTable(4, 4);
                worddoc.Save();
                
            }

            var tvItem = TvSuites.SelectedItem as TreeViewItem;

            if (LoadcomboBoxindex == 0 && tvItem.Items.Count == 0)
            {



                try
                {

                    TbFileNameForExcel.Text = @absolutepath + "\\" + _testProject.TeamProjectName + "\\" + TbFileNameForExcel.Text;

                    var newFile = new FileInfo(TbFileNameForExcel.Text);


                    var template = new FileInfo(Directory.GetCurrentDirectory() + "\\Resources\\TestCaseTemplate.xlsx");
                    using (var xlpackage = new ExcelPackage(newFile, template))
                    {
                        ExcelWorksheet worksheet = xlpackage.Workbook.Worksheets[1];
                        worksheet.Name = (LbSelectTestPlan.SelectedItem as ITestPlan).RootSuite.Title;

                        worksheet.OutLineSummaryBelow = false;
                        //xlpackage.Save();

                        WriteRootSuiteToExcel(rootSuite, worksheet);
                        if (rootSuite.TestSuiteType == TestSuiteType.StaticTestSuite)
                        {
                            //Bug  is here
                            GetTestSuites(rootSuite as IStaticTestSuite, worksheet, xlpackage, null, _testProject, allitemsHeader);
                        }

                        if (rootSuite.TestSuiteType == TestSuiteType.RequirementTestSuite)
                        {
                            GetTestCases(rootSuite as IRequirementTestSuite, worksheet, null, allitemsHeader);
                            xlpackage.Save();
                        }

                        System.Windows.MessageBox.Show("File" + TbFileNameForExcel.Text + " " +  "has been saved to desktop");
                    }
                }
                catch (Exception theException)
                {
                    var errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    System.Windows.MessageBox.Show(errorMessage, "Error");
                }

            }


            else {

                try
                {


                    System.IO.Directory.CreateDirectory(@absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader);

                    TbFileNameForExcel.Text = @absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader + "\\"  + allitemsHeader + ".xlsx";

                    var newFile = new FileInfo(TbFileNameForExcel.Text);


                    var template = new FileInfo(Directory.GetCurrentDirectory() + "\\Resources\\TestCaseTemplate.xlsx");
                    using (var xlpackage = new ExcelPackage(newFile, template))
                    {
                        ExcelWorksheet worksheet = xlpackage.Workbook.Worksheets[1];
                        worksheet.Name = allitemsHeader;

                        worksheet.OutLineSummaryBelow = false;
                        //xlpackage.Save();

                        WriteRootSuiteToExcel(rootSuite, worksheet);
                        if (rootSuite.TestSuiteType == TestSuiteType.StaticTestSuite)
                        {
                            //Bug  is here
                            GetTestSuites(rootSuite as IStaticTestSuite, worksheet, xlpackage, null, _testProject, allitemsHeader);
                        }

                        if (rootSuite.TestSuiteType == TestSuiteType.RequirementTestSuite)
                        {
                            GetTestCases(rootSuite as IRequirementTestSuite, worksheet, null, allitemsHeader);
                            xlpackage.Save();
                        }

                      //  System.Windows.MessageBox.Show("File has been saved at " + TbFileNameForExcel.Text);
                    }
                }
                catch (Exception theException)
                {
                    var errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    System.Windows.MessageBox.Show(errorMessage, "Error");
                }



            }

        }

        void WriteRootSuiteToExcel(ITestSuiteBase testSuite, ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = _testProject + ": " + testSuite.Title;
        }

        void WriteSuiteToExcel(ITestSuiteEntry testSuite, ExcelWorksheet worksheet)
        {
            worksheet.Cells[_i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 1].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
            worksheet.Cells[_i, 1].Style.Font.Bold = true;
            worksheet.Cells[_i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 2].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
            worksheet.Cells[_i, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 3].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
            worksheet.Cells[_i, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 4].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);
            worksheet.Cells[_i, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);

            worksheet.Cells[_i, 1].Value = FindRootSuite(testSuite.TestSuite as IStaticTestSuite, testSuite.Title);
        }

        void WriteSuiteToExcel(IRequirementTestSuite testSuite, ExcelWorksheet worksheet)
        {
            worksheet.Cells[_i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[_i, 1].Style.Font.Bold = true;
            worksheet.Cells[_i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[_i, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 3].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[_i, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 4].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[_i, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[_i, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 6].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            //worksheet.Cells[_i, 1].Value = FindRootSuite(testSuite, testSuite.Title);

        }

        void WriteTestCases(ITestBase testCase, ExcelWorksheet worksheet, DocX worddoc, ITestManagementTeamProject _testProject, string allitemsHeader)
        {

      
            worksheet.Cells[_i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 2].Style.Font.Bold = true;
            worksheet.Cells[_i, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 3].Style.Font.Bold = true;
            worksheet.Cells[_i, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 4].Style.Font.Bold = true;
            worksheet.Cells[_i, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 6].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 7].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 8].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[_i, 9].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            worksheet.Cells[_i, 5].Style.Font.Bold = true;
            worksheet.Cells[_i, 9].Style.Font.Bold = true;
            //worksheet.Cells[_i, 5].Style.Font.Bold = true;

            worksheet.Cells[_i, 1].Value = testCase.Id.ToString();

           

            var testResults = _testProject.TestResults.ByTestId(testCase.Id);

            WorkItemStore Store = _tfs.GetService<WorkItemStore>();


            foreach (ITestCaseResult result in testResults)
            {

                worksheet.Cells[_i, 5].Value = result.Outcome;

                if (result.Outcome.ToString() == "Passed")

                { worksheet.Cells[_i, 5].Style.Fill.BackgroundColor.SetColor(Color.Green); }
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
            "Where [System.TeamProject] = '" + _testProject.WitProject.Name + "'" + "AND ([System.WorkItemType] = 'Bug' ) AND  [System.IterationPath] = '" + testCase.WorkItem.IterationPath + "' AND[System.State] <> 'Closed'");


            WorkItemCollection RequirementQueryResults = Store.Query(
             "SELECT [System.Id],[Microsoft.VSTS.Common.StackRank],[Microsoft.VSTS.Common.Priority], [System.WorkItemType],[System.Title] " +
             "From WorkItems " +
             "Where [System.TeamProject] = '" + _testProject.WitProject.Name + "'" + "AND ([System.WorkItemType] = 'Requirement' ) AND  [System.IterationPath] = '" + testCase.WorkItem.IterationPath + "' AND[System.State] <> 'ANY'");


            /*

            foreach (FieldDefinition disp in BugQueryResults.DisplayFields)
            {
              
                Console.WriteLine(disp.Name);
              
            }

            */


            WorkItem myitem = testCase.WorkItem;

            string _lastvalue = null;
       
            foreach (WorkItemLink wil in myitem.WorkItemLinks)
            {

                foreach (WorkItem disp in BugQueryResults)
                {
                    if (disp.Type.Name == "Bug")
                    {
                        if (disp.Id == wil.TargetId)
                        {
                            _lastvalue += wil.TargetId.ToString() + "-" + disp.Fields["Assigned to"].Value + Environment.NewLine;
                        }

                    }

                }


                foreach (WorkItem disp in RequirementQueryResults)
                {
                    if (disp.Type.Name == "Requirement")
                    {
                        if (disp.Id == wil.TargetId)
                        {
                            worksheet.Cells[_i, 7].Value = disp.Fields["ID"].Value;
                            worksheet.Cells[_i, 8].Value = disp.Fields["Title"].Value;
                        }

                    }

                }


            }

            

            worksheet.Cells[_i, 6].Value = _lastvalue;
            worksheet.Cells[_i, 9].Value = testCase.Priority;
            worksheet.Cells[_i, 2].Value = testCase.Title;
            worksheet.Cells[_i, 3].Value = "Steps:";
            worksheet.Cells[_i, 4].Value = "Expected Results";

            worksheet.Row(_i).OutlineLevel = 1;

            var j = 1;


          
            var wi = Store.GetWorkItem(testCase.Id);
                 


            
            Attachment[] attach = wi.Attachments.Cast<Attachment>().ToArray();
           
            if (attach != null) 
            {

                int i = -1;
                using (System.Net.WebClient request = new System.Net.WebClient())

                    foreach (var attachs in attach)
                    {

                        if (attachs.Extension.ToLower() == ".jpg" || attachs.Extension.ToLower() == ".png")

                        { 

                        i++;
                        string localFilename = @absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader + "\\" + testCase.Id  + "-"  + attachs.Name;

                        if (!System.IO.File.Exists(@localFilename))
                        {
                            System.Uri uri = new System.Uri(localFilename);
                            System.Uri uriForRelativeExcel = new System.Uri(@absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader + "\\" );


                            {
                                request.Credentials = System.Net.CredentialCache.DefaultCredentials;
                                request.DownloadFile(attachs.Uri, localFilename);
                            }
                                worksheet.Cells[_i, 10 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[_i, 10 + i].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                worksheet.Cells[_i, 10 + i].Hyperlink = uriForRelativeExcel.MakeRelativeUri(uri);
                                worksheet.Cells[_i, 10 + i].Value = testCase.Id + "-" + attachs.Name;

                        }

                        else
                        {

                            {
                                request.Credentials = System.Net.CredentialCache.DefaultCredentials;
                                localFilename = @absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader + "\\" + testCase.Id + "-" + attachs.Name;
                                System.Uri uriForRelativeExcel = new System.Uri(@absolutepath + "\\" + _testProject.TeamProjectName + "\\" + allitemsHeader + "\\");
                                request.DownloadFile(attachs.Uri, localFilename);
                                System.Uri uri = new System.Uri(localFilename);
                                    worksheet.Cells[_i, 10 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[_i, 10 + i].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    worksheet.Cells[_i, 10 + i].Hyperlink = uriForRelativeExcel.MakeRelativeUri(uri);
                                    worksheet.Cells[_i, 10 + i].Value = testCase.Id + "-" + attachs.Name;
                            }

                        }
                      }
                 }
                             
            }




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




        static void WriteTestStepsToExcel(ExcelWorksheet worksheet, int i, int j, string testAction, string expectedResult)
        {
            worksheet.Cells[i + j, 3].Value = testAction;
            worksheet.Cells[i + j, 4].Value = expectedResult;

            worksheet.Row(i + j).OutlineLevel = 2;
        }

        private void GetTestCases(IRequirementTestSuite requirementTestSuite, ExcelWorksheet worksheet, DocX worddoc,string allitemsHeader)
        {
            WriteSuiteToExcel(requirementTestSuite, worksheet);
            _i++;
            foreach (var testCase in requirementTestSuite.AllTestCases)
            {
                WriteTestCases(testCase, worksheet, worddoc, _testProject, allitemsHeader);
                _i++;
            }
        }

        private void GetTestSuites(IStaticTestSuite staticTestSuite, ExcelWorksheet worksheet, ExcelPackage xlpackage, DocX worddoc, ITestManagementTeamProject _testProject,string allitemsHeader)
        {
            _i = 3;

            foreach (var suiteEntry in staticTestSuite.Entries.Where(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.TestCase))
               {
                    WriteTestCases(suiteEntry.TestCase, worksheet, worddoc,_testProject, allitemsHeader);
                    _i++;

                }


            foreach (var suiteEntry in staticTestSuite.Entries.Where(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite ||
                                                                                   suiteEntry.EntryType == TestSuiteEntryType.RequirementTestSuite))
            {
                if (suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite)
                {
                    var suite = suiteEntry.TestSuite as IStaticTestSuite;
                    WriteSuiteToExcel(suiteEntry, worksheet);
                    _i++;
                    if (suite.Entries.Count > 0)
                    {
                        GetTestSuites(suite, worksheet, xlpackage, null,_testProject, allitemsHeader);
                    }
                }
                else
                {
                    var suite = suiteEntry.TestSuite as IRequirementTestSuite;
                    GetTestCases(suite, worksheet, worddoc, allitemsHeader);
                }
            }
            xlpackage.Save();

            
        }







        public static string StripTagsCharArray(string source)
        {
            var correctString = source.Replace("&nbsp;", " ");
            var array = new char[correctString.Length];
            var arrayIndex = 0;
            var inside = false;

            foreach (var letter in correctString)
            {
                if (letter == '<')
                {
                    inside = true;
                    continue;
                }
                if (letter == '>')
                {
                    inside = false;
                    continue;
                }
                if (inside) continue;
                array[arrayIndex] = letter;
                arrayIndex++;
            }
            return new string(array, 0, arrayIndex);
        }


        private void BtnOpenFileDialogForExcel_Click(object sender, RoutedEventArgs e)
        {

            var tvItem = TvSuites.SelectedItem as TreeViewItem;

            if (tvItem.Items.Count > 1) { }

            if (TvSuites.SelectedValue != null)
            {
                if (tvItem.Items.Count > 1) { }


                var saveFileDialog1 = new System.Windows.Forms.SaveFileDialog
                {
                    InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString(),
                    Filter = Properties.Resources.MainWindow_BtnOpenFileDialog_Click,
                    FilterIndex = 1
                    
                };

              //  myString.Replace(';', '\n').Replace(',', '\n').Replace('\r', '\n').Replace('\t', '\n').Replace(' ', '\n').Replace("\n\n", "\n");
                saveFileDialog1.FileName = GetTestPlanName(LbSelectTestPlan.SelectedItem as ITestPlan).Replace(':', ' ').Replace ('*' , ' ').Replace( '/',' ');

               // saveFileDialog1.FileName = test_plan_name.Trim(charsToTrim);

                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TbFileNameForExcel.Text= saveFileDialog1.FileName;

                    string mypath =Path.GetFullPath(saveFileDialog1.FileName);
                    absolutepath = System.IO.Path.GetDirectoryName(mypath);
                    int index = mypath.IndexOf(absolutepath) ;
                    
                    string UpdatedExcelFileName = mypath.Substring(absolutepath.Length+1);
                    TbFileNameForExcel.Text = UpdatedExcelFileName;

                }
                else
                {
                    System.Windows.MessageBox.Show("Please choose a valid filename");
                }
            }

            else { System.Windows.MessageBox.Show("Please choose a Test suite"); }

            
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {


            if (TbFileNameForExcel.Text == null || TbFileNameForExcel.Text.Length.Equals(0))
            {
                System.Windows.MessageBox.Show("Please Enter a valid file path");
            }
            else
            {

                //string subPath = null;
                // subPath = @TbFileNameForExcel.Text;// your code goes here

                System.IO.Directory.CreateDirectory(@absolutepath + "\\" + _testProject.TeamProjectName);
            //    string maindirectory = @absolutepath + "\\" + _testProject.TeamProjectName;

               

                _i = 3;




                if (TvSuites.SelectedValue != null)
                {
                    var tvItem = TvSuites.SelectedItem as TreeViewItem;


                    if (tvItem.Items.Count == 0)
                    {

                        _suite = _testProject.TestSuites.Find(Convert.ToInt32(tvItem.Tag));

                        if (_suite != null)
                        {

                            Access_Documents(_suite, _testProject, TbFileNameForExcel.Text);
                        }

                    }


                    else if (tvItem.Items.Count > 1)
                    {
                        List<TreeViewItem> expandedTVI = new List<TreeViewItem>();
                        foreach (TreeViewItem item in tvItem.Items)
                        {
                            expandedTVI.Add(item);
                           
                        }

                        
                        /*
                        foreach (TreeViewItem collectiontest in expandedTVI) {
                            collectiontest.h
                        } 
                        */
                        
                        Parallel.ForEach(expandedTVI, (allitems) =>
                       {
                           
                           this.Dispatcher.BeginInvoke(new Action(() =>
                           
                           Access_Documents(_testProject.TestSuites.Find(Convert.ToInt32(allitems.Tag)), _testProject, allitems.Header.ToString())), null);




                           System.Threading.Thread.Sleep(2);
                       });

                        System.Windows.MessageBox.Show("File" + TbFileNameForExcel.Text + " " + "has been saved to desktop");

                    }


                    else
                    {
                        System.Windows.MessageBox.Show("Please select a test suite");
                    }


                }


            }

        }

        ///Below part is for Requirements generation
        ///


        void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            _tfs = null;
            TbTfs.Text = null;
            var tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            if (tpp.SelectedTeamProjectCollection == null) return;

            _tfs = tpp.SelectedTeamProjectCollection;

            var testService = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));

            _testProject = testService.GetTeamProject(tpp.SelectedProjects[0].Name);

            TbTfs.Text = _tfs.Name + "\\" + _testProject;

            _testPlanCollection = _testProject.TestPlans.Query("Select * from TestPlan");

            ProjectSelected_GetTestPlans();
        }



        private void BtnOpenFileDialog_Click(object sender, RoutedEventArgs e)
        {

            if (!TbTfs.Text.Equals(""))
            {

                System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                openFileDialog1.Title = "Dosya Seçiniz";
                openFileDialog1.InitialDirectory = "C:\\";
                openFileDialog1.ShowDialog();
                TbFileName.Text = openFileDialog1.FileName;

                //OpenAndWorkOnFile(TbFileName.Text);

                OpenXmlPowerTools.GetFileName FileGet = new GetFileName();
                myCollection.AddRange(FileGet.SelectFile(TbFileName.Text));
                Mygrid.IsEnabled = true;
            }

            else
            { MessageBoxResult result = System.Windows.MessageBox.Show("Please Choose Test Project&Plan");}
        }


        private void SaveRequirement_Click(object sender, RoutedEventArgs e)
        {
            // List<string> myCollection = new List<string>();

            if (!TbFileName.Text.Equals(""))
            {

                string name = _testProject.TeamProjectName;
                var _testPlanCollection1 = _testProject.TestPlans.Find(_tvItem);
                string iterationpath = _testPlanCollection1.Iteration;

                WorkItemStore wis = _tfs.GetService<WorkItemStore>();
                Project teamProject = wis.Projects[name];
                WorkItemType workItemType = teamProject.WorkItemTypes["Requirement"];
                
                foreach (var item in Datagriditems)
                {

                    if (item != null)
                    {

                        WorkItem newRequirementWorkitem = new WorkItem(workItemType);
                        newRequirementWorkitem.Title = item.RqName;
                        newRequirementWorkitem.IterationPath = iterationpath;
                        newRequirementWorkitem.Validate();
                        newRequirementWorkitem.Save();


                        // Console.WriteLine(item);


                    }
                }
                MessageBoxResult result = System.Windows.MessageBox.Show("Requirements are Created Succesfully ");
                SaveRequirement.IsEnabled = false;
                // Mygrid.IsEnabled = false;

            }


            else { MessageBoxResult result = System.Windows.MessageBox.Show("Please Upload Requirement Document First"); }

        }

        private void DataGrid_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {

            Datagriditems = new ObservableCollection<MylistElements>();

            foreach (var item in myCollection)
            {

                if (item != null)
                {

                    Datagriditems.Add(new MylistElements(item));
                    //Console.WriteLine(Datagriditems);

                }

            }

            // ... Assign ItemsSource of DataGrid.
            Mygrid.ItemsSource = Datagriditems;


        }

        private void Button_Click_Delete(object sender, RoutedEventArgs e)
        {

            int selected = Mygrid.SelectedIndex;


            Datagriditems.RemoveAt(selected);
            Mygrid.ItemsSource = null;

            Mygrid.ItemsSource = Datagriditems;
            Mygrid.Items.Refresh();

        }

        private void LoadcomboBox_Selected(object sender, RoutedEventArgs e)
        {

            Label1.Visibility = Visibility.Visible;
            DecisionLabel.Visibility = Visibility.Hidden;
            BtnConnect.Visibility = Visibility.Visible;
            TbTfs.Visibility = Visibility.Visible;
            LbSelectTestPlan.Visibility = Visibility.Visible;



            ArrangeVisibilties(LoadcomboBox.SelectedIndex);

        }

        public void ArrangeVisibilties(int Selectedvalue)
        {
                LoadcomboBoxindex = Selectedvalue;

            if (Selectedvalue == 0 || Selectedvalue == 2)
            {
                LoadcomboBox.Visibility = Visibility.Hidden;
                Label3.Visibility = Visibility.Visible;
                TvSuites.Visibility = Visibility.Visible;
                TbFileNameForExcel.Visibility = Visibility.Visible;
                BtnOpenFileDialogForExcel.Visibility = Visibility.Visible;
                BtnGenerate.Visibility = Visibility = Visibility.Visible;
                BtnConnectForExcel.Visibility= Visibility.Visible;

            }

            else
            {
                LoadcomboBox.Visibility = Visibility.Hidden;
                Docpanel.Visibility = Visibility.Visible;
                TvSuites.Visibility = Visibility.Hidden;
                Label2.Visibility = Visibility.Visible;
                TbFileName.Visibility = Visibility.Visible;
                BtnOpenFileDialog.Visibility = Visibility.Visible;
                SaveRequirement.Visibility = Visibility.Visible;
                BtnConnect.Visibility= Visibility.Visible; 

            }

        }

  
    }
}