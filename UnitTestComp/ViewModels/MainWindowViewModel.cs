using DevExpress.Mvvm;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using UnitTestComp.Services.Interface;

namespace UnitTestComp.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ITestService _testService;
        private readonly IRegistryService _registryService;

        public ObservableCollection<string> RecentDllPaths { get; } = new();
        public ObservableCollection<TestCase> AvailableTests { get; } = new();
        public ObservableCollection<TestResult> TestResults { get; } = new();

        private string _selectedDllPath;
        public string SelectedDllPath
        {
            get => _selectedDllPath;
            set
            {
                _selectedDllPath = value;
                OnPropertyChanged(nameof(SelectedDllPath));
                LoadTestsCommand.Execute(null);
            }
        }

        public ICommand BrowseDllCommand { get; }
        public ICommand LoadTestsCommand { get; }
        public ICommand RunTestsCommand { get; }
        public ICommand SaveReportCommand { get; }

        public MainViewModel(ITestService testService, IRegistryService registryService)
        {
            _testService = testService;
            _registryService = registryService;

            BrowseDllCommand = new DelegateCommand(BrowseDll);
            LoadTestsCommand = new DelegateCommand(LoadTests);
            RunTestsCommand = new DelegateCommand(RunTests, CanRunTests);
            SaveReportCommand = new DelegateCommand(SaveReport, CanSaveReport);

            LoadRecentDllPaths();
        }

        private void LoadRecentDllPaths()
        {
            var paths = _registryService.GetRecentDllPaths();
            RecentDllPaths.Clear();
            foreach (var path in paths)
            {
                RecentDllPaths.Add(path);
            }
        }

        private void BrowseDll()
        {
            var path = _testService.BrowseForDll();
            if (!string.IsNullOrEmpty(path))
            {
                SelectedDllPath = path;
                _registryService.SaveDllPath(path);
            }
        }

        private async void LoadTests()
        {
            var tests = await _testService.LoadTestsAsync(SelectedDllPath);
            AvailableTests.Clear();
            foreach (var test in tests)
            {
                AvailableTests.Add(test);
            }
        }

        private async void RunTests()
        {
            var results = await _testService.RunTestsAsync(AvailableTests.ToList());
            TestResults.Clear();
            foreach (var result in results)
            {
                TestResults.Add(result);
            }
        }

        private void SaveReport()
        {
            _testService.CreateWordReport(TestResults.ToList());
        }

        private bool CanRunTests() => AvailableTests.Any();
        private bool CanSaveReport() => TestResults.Any();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
