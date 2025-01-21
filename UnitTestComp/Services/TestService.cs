using DevExpress.Xpf.Grid;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.TestPlatform.VsTestConsole.TranslationLayer;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using UnitTestComp.Handlers;
using UnitTestComp.Services.Interface;

namespace UnitTestComp.Services
{
    public class TestService : ITestService
    {
        private readonly VsTestConsoleWrapper _vstestConsoleWrapper;
        private readonly IRegistryService _registryService;

        public TestService(string vstestConsolePath, IRegistryService registryService)
        {
            _vstestConsoleWrapper = new VsTestConsoleWrapper(vstestConsolePath);
            _vstestConsoleWrapper.StartSession();
            _registryService = registryService;
        }

        public string BrowseForDll()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "DLL files (*.dll)|*.dll",
                Title = "Select Unit Test DLL"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string selectedPath = openFileDialog.FileName;
                _registryService.GetRecentDllPaths();
                return selectedPath;
            }

            return string.Empty;
        }

        public async Task<List<TestCase>> LoadTestsAsync(string dllPath)
        {
            var discoveryEventHandler = new DiscoveryEventHandler();
            bool discoveryCompleted = false;

            if (!File.Exists(dllPath))
            {
                throw new FileNotFoundException("DLL file not found.");
            }

            try
            {
                discoveryCompleted = await Task.Run(() =>
                {
                    _vstestConsoleWrapper.DiscoverTests(new List<string> { dllPath }, null, discoveryEventHandler);
                    return true;
                }).WaitAsync(TimeSpan.FromMinutes(10));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error during test discovery: {ex.Message}");
            }

            if (!discoveryCompleted)
            {
                throw new TimeoutException("Test discovery timed out.");
            }

            return discoveryEventHandler.GetDiscoveredTestCases().ToList();
        }

        public async Task<List<TestResult>> RunTestsAsync(List<TestCase> testCases)
        {
            if (testCases == null || testCases.Count == 0)
            {
                throw new ArgumentException("No tests selected for execution.");
            }

            var runEventHandler = new RunEventHandler();
            bool testRunCompleted = false;

            try
            {
                testRunCompleted = await Task.Run(() =>
                {
                    _vstestConsoleWrapper.RunTests(testCases, null, runEventHandler);
                    return true;
                }).WaitAsync(TimeSpan.FromMinutes(10));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error during test execution: {ex.Message}");
            }

            if (!testRunCompleted)
            {
                throw new Exception("Test execution failed.");
            }

            return runEventHandler.TestResults.ToList();
        }

        public void CreateWordReport(List<TestResult> results)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Documents (*.docx)|*.docx";
            saveFileDialog.Title = "Save Report As...";
            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                if (string.IsNullOrEmpty(filePath))
                    return;

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    Body body = mainPart.Document.Body;

                    Paragraph title = new Paragraph(new Run(new Text("Test Results Report")));
                    body.Append(title);

                    Table table = new Table();

                    TableProperties tableProperties = new TableProperties(
                        new TableWidth { Type = TableWidthUnitValues.Auto },
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 6 },
                            new BottomBorder { Val = BorderValues.Single, Size = 6 },
                            new LeftBorder { Val = BorderValues.Single, Size = 6 },
                            new RightBorder { Val = BorderValues.Single, Size = 6 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                        )
                    );
                    table.Append(tableProperties);

                    // Add table headers
                    TableRow headerRow = new TableRow();
                    headerRow.Append(
                        new TableCell(new Paragraph(new Run(new Text("Test Class")))),
                        new TableCell(new Paragraph(new Run(new Text("Method Name")))),
                        new TableCell(new Paragraph(new Run(new Text("Outcome")))),
                        new TableCell(new Paragraph(new Run(new Text("Duration"))))
                    );
                    table.Append(headerRow);

                    // Add test results to the table
                    foreach (var result in results)
                    {
                        TableRow row = new TableRow();
                        row.Append(
                            new TableCell(new Paragraph(new Run(new Text(GetFirstWordWithIndexOf(result.TestCase.DisplayName))))),
                            new TableCell(new Paragraph(new Run(new Text(RemoveFirstPart(result.TestCase.DisplayName))))),
                            new TableCell(new Paragraph(new Run(new Text(result.Outcome.ToString())))),
                            new TableCell(new Paragraph(new Run(new Text(result.Duration.ToString()))))
                        );
                        table.Append(row);
                    }

                    body.Append(table);

                    var failedTests = results.Where(r => r.Outcome != TestOutcome.Passed).ToList();
                    if (failedTests.Any())
                    {
                        body.Append(new Paragraph(new Run(new Text("\nReasons for failed tests:"))));
                        foreach (var failedTest in failedTests)
                        {
                            var reason = new Paragraph();
                            reason.Append(new Run(new Text($"Test: {failedTest.TestCase.FullyQualifiedName}")));
                            reason.Append(new Run(new Break()));
                            reason.Append(new Run(new Text($"Outcome: {failedTest.Outcome}")));
                            reason.Append(new Run(new Break()));
                            reason.Append(new Run(new Text($"Error Message: {failedTest.ErrorMessage ?? "No error message"}")));
                            reason.Append(new Run(new Break()));
                            body.Append(reason);
                        }
                    }
                    else
                    {
                        body.Append(new Paragraph(new Run(new Text("\nAll tests passed successfully."))));
                    }

                    MessageBox.Show($"Tests completed, report saved at {filePath}");
                }
            }
            else
            {
                MessageBox.Show("Save process was canceled.");
            }
        }

        private string RemoveFirstPart(string input)
        {
            int firstDotIndex = input.IndexOf('.');
            if (firstDotIndex >= 0)
            {
                return input.Substring(firstDotIndex + 1);
            }

            return input;
        }

        private string GetFirstWordWithIndexOf(string input)
        {
            int firstDotIndex = input.IndexOf('.');

            return firstDotIndex >= 0 ? input.Substring(0, firstDotIndex) : input;
        }
    }

}
