using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Logging;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System;
using System.Collections.Generic;

namespace UnitTestComp.Handlers
{
    public class DiscoveryEventHandler : ITestDiscoveryEventsHandler
    {
        private readonly List<TestCase> _discoveredTestCases;

        public DiscoveryEventHandler()
        {
            _discoveredTestCases = new List<TestCase>();
        }

        public void HandleDiscoveredTests(IEnumerable<TestCase> discoveredTestCases)
        {
            foreach (var testCase in discoveredTestCases)
            {
                Console.WriteLine($"Discovered test: {testCase.FullyQualifiedName}");
                _discoveredTestCases.Add(testCase);  // Добавляем каждый тест в список
            }
        }

        public void HandleDiscoveryComplete(long totalTests, IEnumerable<TestCase> lastDiscoveredTests, bool isAborted)
        {
            Console.WriteLine($"Discovery completed. Total tests: {totalTests}");
        }

        public void HandleLogMessage(TestMessageLevel level, string message)
        {
            Console.WriteLine($"[{level}] {message}");
        }

        public void HandleRawMessage(string rawMessage)
        {
            Console.WriteLine(rawMessage);
        }

        // Возвращаем список найденных тестов
        public List<TestCase> GetDiscoveredTestCases() => _discoveredTestCases;
    }

}
