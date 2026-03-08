using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Bug condition exploration tests for GUI inspection mode fix.
    /// 
    /// CRITICAL: These tests are EXPECTED TO FAIL on unfixed code.
    /// Test failure confirms the bug exists.
    /// 
    /// These tests encode the EXPECTED (correct) behavior.
    /// When the bug is fixed, these tests will pass.
    /// 
    /// **Validates: Requirements 2.1, 2.2, 2.3**
    /// </summary>
    [TestFixture]
    public class GUIInspectionModeBugExplorationTests
    {
        /// <summary>
        /// Property 1: Fault Condition - InspectSheets Project Configuration
        /// 
        /// This test verifies that the InspectSheets project is configured correctly
        /// to NOT create an executable entry point that conflicts with the main GUI application.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL because:
        /// - InspectSheets/InspectSheets.csproj has OutputType=Exe
        /// - This creates a conflicting Main() entry point
        /// - The build system cannot determine which entry point to use
        /// - Result: "error CS0017: Nel programma è definito più di un punto di ingresso"
        /// 
        /// When the bug is fixed, this test will PASS.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_InspectSheetsProject_ShouldNotHaveExecutableOutputType()
        {
            // Arrange
            var inspectSheetsCsprojPath = Path.Combine("InspectSheets", "InspectSheets.csproj");
            
            Console.WriteLine("=== ANALYZING INSPECTSHEETS PROJECT CONFIGURATION ===");
            Console.WriteLine($"Project file: {inspectSheetsCsprojPath}");
            
            Assert.That(File.Exists(inspectSheetsCsprojPath), Is.True,
                $"InspectSheets project file not found at: {inspectSheetsCsprojPath}");

            // Act - Read and parse the project file
            var projectXml = XDocument.Load(inspectSheetsCsprojPath);
            var outputTypeElement = projectXml.Descendants("OutputType").FirstOrDefault();
            
            Console.WriteLine($"\nCurrent OutputType: {outputTypeElement?.Value ?? "(not specified)"}");
            
            // Document the bug condition
            if (outputTypeElement != null && outputTypeElement.Value == "Exe")
            {
                Console.WriteLine("\n=== BUG CONDITION DETECTED ===");
                Console.WriteLine("COUNTEREXAMPLE: InspectSheets.csproj has OutputType=Exe");
                Console.WriteLine("This creates a conflicting entry point with the main Program.cs");
                Console.WriteLine("\nEvidence of the bug:");
                Console.WriteLine("1. InspectSheets/Program.cs contains a Main() method");
                Console.WriteLine("2. Root Program.cs also contains a Main() method");
                Console.WriteLine("3. Build fails with: 'error CS0017: Nel programma è definito più di un punto di ingresso'");
                Console.WriteLine("4. When published, the wrong entry point may be selected");
                Console.WriteLine("\nExpected behavior:");
                Console.WriteLine("- InspectSheets should have OutputType=Library (not Exe)");
                Console.WriteLine("- This prevents it from creating an executable entry point");
                Console.WriteLine("- Only the main Program.cs entry point should be used");
            }

            // Assert - Verify CORRECT behavior (will fail on unfixed code)
            Assert.That(outputTypeElement, Is.Not.Null,
                "COUNTEREXAMPLE: OutputType element not found in InspectSheets.csproj. " +
                "The project should explicitly specify OutputType=Library to prevent creating an executable.");

            Assert.That(outputTypeElement.Value, Is.Not.EqualTo("Exe"),
                "COUNTEREXAMPLE: InspectSheets/InspectSheets.csproj has OutputType=Exe, creating a conflicting entry point. " +
                "This causes the build error 'error CS0017: Nel programma è definito più di un punto di ingresso' " +
                "and may cause the wrong entry point to be executed when the application is published.\n\n" +
                $"Current value: {outputTypeElement.Value}\n" +
                "Expected value: Library (or any value other than 'Exe')\n\n" +
                "Root cause: The InspectSheets project is a utility for debugging and should not create an executable. " +
                "It should be compiled as a library to prevent entry point conflicts.\n\n" +
                "Impact: When published, the application may run InspectSheets/Program.cs instead of the main GUI Program.cs, " +
                "displaying 'Inspecting file: ...' output instead of opening the MainForm GUI window.");

            Console.WriteLine("\n=== TEST PASSED ===");
            Console.WriteLine($"InspectSheets.csproj has correct OutputType: {outputTypeElement.Value}");
            Console.WriteLine("No conflicting entry point will be created.");
        }

        /// <summary>
        /// Property 1: Fault Condition - Build Succeeds Without Entry Point Conflicts
        /// 
        /// This test verifies that the project can be built successfully without
        /// multiple entry point conflicts.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL because:
        /// - Build fails with error CS0017: "Nel programma è definito più di un punto di ingresso"
        /// - Both Program.cs and InspectSheets/Program.cs have Main() methods
        /// - The compiler cannot determine which entry point to use
        /// 
        /// When the bug is fixed, this test will PASS.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_ProjectBuild_ShouldSucceedWithoutEntryPointConflicts()
        {
            // Arrange
            Console.WriteLine("=== ATTEMPTING TO BUILD PROJECT ===");
            Console.WriteLine("This test verifies that the project builds without entry point conflicts.");
            
            // Act - Try to build the project
            var buildResult = RunProcess("dotnet", "build --no-incremental", timeoutMs: 60000);
            
            Console.WriteLine($"\nBuild exit code: {buildResult.ExitCode}");
            Console.WriteLine($"Build output:\n{buildResult.Output}");
            
            if (!string.IsNullOrWhiteSpace(buildResult.Error))
            {
                Console.WriteLine($"Build errors:\n{buildResult.Error}");
            }

            // Document the bug if it exists
            if (buildResult.Output.Contains("CS0017") || buildResult.Output.Contains("più di un punto di ingresso"))
            {
                Console.WriteLine("\n=== BUG CONDITION DETECTED ===");
                Console.WriteLine("COUNTEREXAMPLE: Build failed with multiple entry point error (CS0017)");
                Console.WriteLine("\nThis confirms the bug:");
                Console.WriteLine("1. InspectSheets/InspectSheets.csproj has OutputType=Exe");
                Console.WriteLine("2. This creates a Main() entry point in InspectSheets/Program.cs");
                Console.WriteLine("3. The root Program.cs also has a Main() entry point");
                Console.WriteLine("4. The compiler cannot determine which entry point to use");
                Console.WriteLine("\nExpected behavior after fix:");
                Console.WriteLine("- InspectSheets should have OutputType=Library");
                Console.WriteLine("- Build should succeed with only one entry point (root Program.cs)");
                Console.WriteLine("- Application should launch GUI when executed");
            }

            // Assert - Verify CORRECT behavior (will fail on unfixed code)
            Assert.That(buildResult.ExitCode, Is.EqualTo(0),
                "COUNTEREXAMPLE: Build failed. This is expected on unfixed code due to multiple entry points.\n" +
                "Error CS0017: 'Nel programma è definito più di un punto di ingresso'\n\n" +
                "Root cause: InspectSheets/InspectSheets.csproj has OutputType=Exe, creating a conflicting Main() method.\n" +
                "Both Program.cs and InspectSheets/Program.cs define entry points, causing the build to fail.\n\n" +
                "Expected behavior: Build should succeed with only the main Program.cs entry point.");

            Assert.That(buildResult.Output, Does.Not.Contain("CS0017"),
                "COUNTEREXAMPLE: Build output contains error CS0017 (multiple entry points defined).\n" +
                "This confirms the bug: InspectSheets creates a conflicting entry point.");

            Assert.That(buildResult.Output, Does.Not.Contain("più di un punto di ingresso"),
                "COUNTEREXAMPLE: Build output contains 'più di un punto di ingresso' (more than one entry point).\n" +
                "This confirms the bug: InspectSheets creates a conflicting entry point.");

            Console.WriteLine("\n=== TEST PASSED ===");
            Console.WriteLine("Build succeeded without entry point conflicts.");
            Console.WriteLine("Only the main Program.cs entry point is being used.");
        }

        /// <summary>
        /// Runs a process and captures its output.
        /// </summary>
        private ProcessResult RunProcess(string fileName, string arguments, int timeoutMs)
        {
            var result = new ProcessResult();
            var outputBuilder = new System.Text.StringBuilder();
            var errorBuilder = new System.Text.StringBuilder();

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = Directory.GetCurrentDirectory()
                };

                using (var process = new Process { StartInfo = startInfo })
                {
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (e.Data != null)
                        {
                            outputBuilder.AppendLine(e.Data);
                        }
                    };

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (e.Data != null)
                        {
                            errorBuilder.AppendLine(e.Data);
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    bool exited = process.WaitForExit(timeoutMs);

                    if (!exited)
                    {
                        // Process didn't exit within timeout
                        try
                        {
                            process.Kill();
                        }
                        catch
                        {
                            // Ignore errors when killing the process
                        }
                        
                        result.Success = false;
                        result.ExitCode = -1;
                        result.Error = "Process timed out";
                    }
                    else
                    {
                        // Process exited - capture exit code
                        result.Success = process.ExitCode == 0;
                        result.ExitCode = process.ExitCode;
                    }
                }

                result.Output = outputBuilder.ToString();
                result.Error = errorBuilder.ToString();
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
            }

            return result;
        }

        private class ProcessResult
        {
            public bool Success { get; set; }
            public int ExitCode { get; set; }
            public string Output { get; set; } = "";
            public string Error { get; set; } = "";
        }
    }
}
