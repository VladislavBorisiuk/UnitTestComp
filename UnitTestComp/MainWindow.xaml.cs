using DevExpress.Xpf.Core;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Grid;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.TestPlatform.VsTestConsole.TranslationLayer;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Logging;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UnitTestComp.Handlers;
using static DevExpress.Xpo.Helpers.CannotLoadObjectsHelper;

namespace UnitTestComp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ThemedWindow
    {
        private const string RegistryKeyPath = @"SOFTWARE\UnitTestComp";
        private const string RecentDllPathsValueName = "RecentDllPaths";
        private RunEventHandler runEventHandler;
        private VsTestConsoleWrapper _vstestConsoleWrapper;

        public MainWindow()
        {
            InitializeComponent();
            InitializeTestPlatform();
            LoadRecentDllPaths(); // Загружаем последние пути при запуске
        }

        private void LoadRecentDllPaths()
        {
            using (RegistryKey userKey = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (userKey != null)
                {
                    string recentPaths = userKey.GetValue(RecentDllPathsValueName) as string;
                    if (!string.IsNullOrEmpty(recentPaths))
                    {
                        var paths = recentPaths.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        RecentDllsComboBox.Items.Clear();
                        foreach (var path in paths)
                        {
                            RecentDllsComboBox.Items.Add(path);
                        }
                    }
                }
            }
        }

        private void SaveDllPathToRegistry(string newPath)
        {
            using (RegistryKey userKey = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                string recentPaths = userKey.GetValue(RecentDllPathsValueName) as string;
                var paths = string.IsNullOrEmpty(recentPaths)
                    ? new List<string>()
                    : recentPaths.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                // Добавляем новый путь, если его ещё нет в списке
                if (!paths.Contains(newPath))
                {
                    paths.Insert(0, newPath);
                    if (paths.Count > 5) // Оставляем только последние 5 путей
                    {
                        paths = paths.Take(5).ToList();
                    }
                }

                // Сохраняем обратно в реестр
                userKey.SetValue(RecentDllPathsValueName, string.Join("|", paths));
            }

            LoadRecentDllPaths(); // Обновляем список путей в интерфейсе
        }

        private async void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "DLL files (*.dll)|*.dll",
                Title = "Select Unit Test DLL"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string selectedPath = openFileDialog.FileName;
                RecentDllsComboBox.SelectedItem = selectedPath;
                SaveDllPathToRegistry(selectedPath); // Сохраняем выбранный путь

                await LoadTests(selectedPath); // Загружаем тесты из выбранного DLL файла
            }
        }

        private void RecentDllsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RecentDllsComboBox.SelectedItem is string selectedPath)
            {
                RecentDllsComboBox.SelectedItem = selectedPath;
                LoadTests(selectedPath);
            }
        }

        private void InitializeTestPlatform()
        {
            string vstestConsolePath = GetVstestConsolePathFromRegistry();

            if (string.IsNullOrEmpty(vstestConsolePath) || !File.Exists(vstestConsolePath))
            {
                MessageBox.Show("Путь к vstest.console.exe не найден или недействителен. Укажите путь вручную.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);

                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Executable files (*.exe)|*.exe",
                    Title = "Укажите путь до vstest.console.exe"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    vstestConsolePath = openFileDialog.FileName;

                    if (File.Exists(vstestConsolePath) && Path.GetFileName(vstestConsolePath).Equals("vstest.console.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        SaveVstestConsolePathToRegistry(vstestConsolePath);
                    }
                    else
                    {
                        MessageBox.Show("Указан недействительный путь к vstest.console.exe.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Путь к vstest.console.exe не указан.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            _vstestConsoleWrapper = new VsTestConsoleWrapper(vstestConsolePath);
            _vstestConsoleWrapper.StartSession();
        }

        private string GetVstestConsolePathFromRegistry()
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\VisualStudio\17.0"))
            {
                if (key != null)
                {
                    var installPath = key.GetValue("InstallDir") as string;
                    if (!string.IsNullOrEmpty(installPath))
                    {
                        return Path.Combine(installPath, "Common7", "IDE", "Extensions", "TestPlatform", "vstest.console.exe");
                    }
                }
            }

            using (RegistryKey userKey = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (userKey != null)
                {
                    return userKey.GetValue("VstestConsolePath") as string;
                }
            }

            return null;
        }

        private void SaveVstestConsolePathToRegistry(string path)
        {
            using (RegistryKey userKey = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                userKey.SetValue("VstestConsolePath", path);
            }

            MessageBox.Show("Путь к vstest.console.exe успешно сохранён в реестр.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private void CreateWordReport(IEnumerable<TestResult> results)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Documents (*.docx)|*.docx";
            saveFileDialog.Title = "Сохранить отчет как...";
            var dialogResult = saveFileDialog.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value)
            {
                string filePath = saveFileDialog.FileName;
                if (string.IsNullOrEmpty(filePath))
                    return;
                ShowWaitIndicator("Генерируемый отчет...");

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    Body body = mainPart.Document.Body;
                    Paragraph title = new Paragraph(new Run(new Text("Отчет о результатах теста")));
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

                    TableRow headerRow = new TableRow();
                    headerRow.Append(
                        new TableCell(new Paragraph(new Run(new Text("Класс теста")))),
                        new TableCell(new Paragraph(new Run(new Text("Имя метода")))),
                        new TableCell(new Paragraph(new Run(new Text("Результат")))),
                        new TableCell(new Paragraph(new Run(new Text("Продолжительность"))))
                    );
                    table.Append(headerRow);

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
                        body.Append(new Paragraph(new Run(new Text("\nПричины неудачных тестов:"))));
                        foreach (var failedTest in failedTests)
                        {
                            var reason = new Paragraph();
                            reason.Append(new Run(new Text("Тест: " + failedTest.TestCase.FullyQualifiedName)));
                            reason.Append(new Run(new Break()));
                            reason.Append(new Run(new Text("Результат: " + failedTest.Outcome)));
                            reason.Append(new Run(new Break()));
                            reason.Append(new Run(new Text("Сообщение об ошибке: " + (failedTest.ErrorMessage ?? "Нет сообщения об ошибке"))));
                            reason.Append(new Run(new Break())); 
                            reason.Append(new Run(new Break())); 

                            body.Append(reason);
                        }
                    }
                    else
                    {
                        body.Append(new Paragraph(new Run(new Text("\nВсе тесты пройдены успешно."))));
                    }

                    SplashScreenManager.CloseAll();
                    MessageBox.Show($"Тесты выполнены, отчет сохранен {filePath}");
                }
            }
            else
            {
                MessageBox.Show("Процесс сохранения был отменен.");
            }

        }


        private void DisplayTestResults(IEnumerable<TestResult> results)
        {
            ResultGrid.ItemsSource = results;
            foreach (var result in results)
            {
                result.TestCase.DisplayName = RemoveFirstPart(result.TestCase.DisplayName);
            }
            ResultGrid.RefreshData();
            CreateReportButton.IsEnabled = true;
        }



        private async Task LoadTests(string testDllPath)
        {
            ShowWaitIndicator("Загружаем тесты...");
            if (!File.Exists(testDllPath))
            {
                MessageBox.Show("DLL file not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var discoveryEventHandler = new DiscoveryEventHandler();
            try
            {
                bool discoveryCompleted = await Task.Run(() =>
                {
                    _vstestConsoleWrapper.DiscoverTests(new List<string> { testDllPath }, null, discoveryEventHandler);
                    return true;
                }).WaitAsync(TimeSpan.FromSeconds(30));

                if (!discoveryCompleted)
                {
                    MessageBox.Show("Discovery timed out.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var discoveredTests = discoveryEventHandler.GetDiscoveredTestCases();

                TestListBox.Items.Clear();
                foreach (var testCase in discoveredTests)
                {
                    TestListBox.Items.Add(new ListBoxEditItem { Content = testCase.FullyQualifiedName, Tag = testCase });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            SplashScreenManager.CloseAll();
        }
        private void CreateReportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CreateWordReport(runEventHandler.TestResults);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

            private async void RunTestsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowWaitIndicator("Выполняем тесты...");
            try
            {
                var selectedTests = TestListBox.SelectedItems.Cast<ListBoxItem>()
                    .Select(item => (TestCase)item.Tag).ToList();

                if (selectedTests.Count == 0)
                {
                    MessageBox.Show("Please select at least one test to run.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                await RunSelectedTestsAsync(selectedTests); 
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SplashScreenManager.CloseAll();
            }
        }

        private async Task RunSelectedTestsAsync(List<TestCase> selectedTests)
        {
            runEventHandler = new RunEventHandler();

            string testDllPath = RecentDllsComboBox.Text;

            var selectedTestNames = selectedTests.Select(test => test.FullyQualifiedName).ToList();
            string testsArgument = string.Join(",", selectedTestNames);

            try
            {
                bool testRunCompleted = await Task.Run(() =>
                {
                    try
                    {
                        _vstestConsoleWrapper.RunTests(selectedTests, null, runEventHandler);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Test execution failed: {ex.Message}");
                        return false;
                    }
                });

                if (!testRunCompleted)
                {
                    MessageBox.Show("Test execution failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                DisplayTestResults(runEventHandler.TestResults);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void Window_Closed(object sender, EventArgs e)
        {
            _vstestConsoleWrapper?.EndSession();
        }

        private void ShowWaitIndicator(string status)
        {
            SplashScreenManager.CreateWaitIndicator(new DevExpress.Mvvm.DXSplashScreenViewModel { Status = status }).Show();
        }

        string RemoveFirstPart(string input)
        {
            int firstDotIndex = input.IndexOf('.');
            if (firstDotIndex >= 0)
            {
                return input.Substring(firstDotIndex + 1);
            }

            return input;
        }

        string GetFirstWordWithIndexOf(string input)
        {
            int firstDotIndex = input.IndexOf('.');

            return firstDotIndex >= 0 ? input.Substring(0, firstDotIndex) : input;
        }
    }
}

