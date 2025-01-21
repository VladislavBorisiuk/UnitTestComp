using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Logging;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System;
using System.Collections.Generic;

namespace UnitTestComp.Handlers
{
    public class RunEventHandler : ITestRunEventsHandler
    {
        public List<TestResult> TestResults { get; } = new List<TestResult>();

        public void HandleLogMessage(TestMessageLevel level, string message) => Console.WriteLine($"[{level}] {message}");

        public void HandleTestRunStatsChange(TestRunChangedEventArgs statsChange)
        {
            if (statsChange.NewTestResults != null)
                TestResults.AddRange(statsChange.NewTestResults);
        }

        public void HandleTestRunComplete(TestRunCompleteEventArgs completeArgs, TestRunChangedEventArgs lastChunk, ICollection<AttachmentSet> runContextAttachments, ICollection<string> executorUris)
        {
            if (lastChunk?.NewTestResults != null)
                TestResults.AddRange(lastChunk.NewTestResults);
        }

        public void HandleRawMessage(string rawMessage) => Console.WriteLine(rawMessage);
        public int LaunchProcessWithDebuggerAttached(TestProcessStartInfo testProcessStartInfo) => -1;

    }
}
