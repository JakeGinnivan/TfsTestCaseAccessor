using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace TfsTestCaseAccessor
{
    public static class TestCases
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="serverName">The tfs server for example http://server:8080/tfs/Collection/ </param>
        /// <param name="projectName">The project name</param>
        /// <param name="assemblyName">Can be fully qualified ie MyTests.dll</param>
        /// <param name="testMethod">The fully qualified test method</param>
        /// <returns></returns>
        public static ITestCase GetTestCase(string serverName, string projectName, string assemblyName = null, string testMethod = null)
        {
            var tfs = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(serverName));
            var testService = tfs.GetService<ITestManagementService>();
            var project = testService.GetTeamProject(projectName);

            // Query the test case.
            const string interestedFields = "[System.Id], [System.Title]"; // and more
            var testCaseName = testMethod ?? GetTestCaseName();
            var storageName = Path.GetFileName(assemblyName ?? Assembly.GetCallingAssembly().CodeBase);
            string query = string.Format(CultureInfo.InvariantCulture,
                "SELECT {0} FROM WorkItems WHERE Microsoft.VSTS.TCM.AutomatedTestName = '{1}'" +
                "AND Microsoft.VSTS.TCM.AutomatedTestStorage = '{2}'",
                interestedFields, testCaseName, storageName);

            var testCases = project.TestCases.Query(query).ToArray();
            return testCases.Length == 1 ? testCases.First() : null;
        }

        private static string GetTestCaseName()
        {
            var stackFrame = new StackFrame(2);
            var methodBase = stackFrame.GetMethod();
            var testCaseName = methodBase.DeclaringType + "." + methodBase.Name;
            return testCaseName;
        }
    }
}
