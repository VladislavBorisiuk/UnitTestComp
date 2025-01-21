using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace UnitTestComp.Services.Interface
{
    public interface ITestService
    {
        string BrowseForDll();
        void CreateWordReport(List<TestResult> results);
        Task<List<TestCase>> LoadTestsAsync(string dllPath);
        Task<List<TestResult>> RunTestsAsync(List<TestCase> testCases);
    }
}