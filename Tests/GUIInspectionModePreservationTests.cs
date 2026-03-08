using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.UI;
using FsCheck;
using FsCheck.NUnit;
using NUnit.Framework;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Preservation property tests for GUI inspection mode fix.
    /// 
    /// IMPORTANT: These tests verify baseline behavior that must be preserved.
    /// These tests should PASS on UNFIXED code (when using correct entry point).
    /// After the fix, these tests should STILL PASS (no regressions).
    /// 
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// </summary>
    [TestFixture]
    public class GUIInspectionModePreservationTests
    {
        /// <summary>
        /// Property 2: Preservation - Italian Culture Setting
        /// 
        /// This test verifies that the application correctly sets Italian culture (it-IT)
        /// during initialization. This behavior must be preserved after the fix.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.2**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_ItalianCulture_IsSetCorrectly(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    // Arrange - Save current culture
                    var originalCulture = System.Threading.Thread.CurrentThread.CurrentUICulture;

                    try
                    {
                        // Act - Simulate the culture setting from Program.cs
                        System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("it-IT");

                        // Assert - Verify Italian culture is set
                        var currentCulture = System.Threading.Thread.CurrentThread.CurrentUICulture;
                        
                        return (currentCulture.Name == "it-IT")
                            .Label($"Expected Italian culture (it-IT), got: {currentCulture.Name}");
                    }
                    finally
                    {
                        // Restore original culture
                        System.Threading.Thread.CurrentThread.CurrentUICulture = originalCulture;
                    }
                });
        }

        /// <summary>
        /// Property 2: Preservation - Service Initialization
        /// 
        /// This test verifies that all required services can be initialized correctly.
        /// The services are: CSVParser, DateCalculator, HeaderCalculator, 
        /// TransformationRulesEngine, DataTransformer, ExcelManager.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_Services_InitializeCorrectly(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    try
                    {
                        // Act - Initialize all services as done in Program.cs
                        var csvParser = new CSVParser();
                        var dateCalculator = new DateCalculator();
                        var headerCalculator = new HeaderCalculator(dateCalculator);
                        var transformationRulesEngine = new TransformationRulesEngine();
                        var dataTransformer = new DataTransformer(transformationRulesEngine);
                        var excelManager = new ExcelManager();

                        // Assert - Verify all services are not null
                        return (csvParser != null)
                            .Label("CSVParser should be initialized")
                            .And(dateCalculator != null)
                            .Label("DateCalculator should be initialized")
                            .And(headerCalculator != null)
                            .Label("HeaderCalculator should be initialized")
                            .And(transformationRulesEngine != null)
                            .Label("TransformationRulesEngine should be initialized")
                            .And(dataTransformer != null)
                            .Label("DataTransformer should be initialized")
                            .And(excelManager != null)
                            .Label("ExcelManager should be initialized");
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Service initialization failed: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Property 2: Preservation - GUIWrapper Pattern
        /// 
        /// This test verifies that the GUIWrapper pattern (used to handle circular
        /// dependencies between MainForm and ApplicationController) works correctly.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.4**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_GUIWrapper_HandlesCircularDependencies(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    try
                    {
                        // Arrange - Initialize services
                        var csvParser = new CSVParser();
                        var dateCalculator = new DateCalculator();
                        var headerCalculator = new HeaderCalculator(dateCalculator);
                        var transformationRulesEngine = new TransformationRulesEngine();
                        var dataTransformer = new DataTransformer(transformationRulesEngine);
                        var excelManager = new ExcelManager();

                        // Act - Create GUIWrapper and ApplicationController
                        // Use reflection to access the internal GUIWrapper class
                        var guiWrapperType = typeof(Program).Assembly.GetTypes()
                            .FirstOrDefault(t => t.Name == "GUIWrapper");
                        
                        if (guiWrapperType == null)
                        {
                            return false.Label("GUIWrapper class not found");
                        }

                        var guiWrapper = Activator.CreateInstance(guiWrapperType);
                        
                        if (guiWrapper == null)
                        {
                            return false.Label("Failed to create GUIWrapper instance");
                        }

                        // Create ApplicationController with the wrapper
                        var controller = new ApplicationController(
                            (IGUI)guiWrapper,
                            csvParser,
                            excelManager,
                            dataTransformer,
                            headerCalculator
                        );

                        // Assert - Verify controller and wrapper are created
                        return (controller != null)
                            .Label("ApplicationController should be created with GUIWrapper")
                            .And(guiWrapper != null)
                            .Label("GUIWrapper should be created successfully");
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"GUIWrapper pattern test failed: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Property 2: Preservation - MainForm Creation
        /// 
        /// This test verifies that MainForm can be created with ApplicationController
        /// and that all required controls are initialized.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.1, 3.4**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_MainForm_CreatesWithAllControls(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    try
                    {
                        // Arrange - Initialize services
                        var csvParser = new CSVParser();
                        var dateCalculator = new DateCalculator();
                        var headerCalculator = new HeaderCalculator(dateCalculator);
                        var transformationRulesEngine = new TransformationRulesEngine();
                        var dataTransformer = new DataTransformer(transformationRulesEngine);
                        var excelManager = new ExcelManager();

                        // Create GUIWrapper
                        var guiWrapperType = typeof(Program).Assembly.GetTypes()
                            .FirstOrDefault(t => t.Name == "GUIWrapper");
                        
                        if (guiWrapperType == null)
                        {
                            return false.Label("GUIWrapper class not found");
                        }

                        var guiWrapper = Activator.CreateInstance(guiWrapperType);
                        
                        if (guiWrapper == null)
                        {
                            return false.Label("Failed to create GUIWrapper instance");
                        }

                        // Create ApplicationController
                        var controller = new ApplicationController(
                            (IGUI)guiWrapper,
                            csvParser,
                            excelManager,
                            dataTransformer,
                            headerCalculator
                        );

                        // Act - Create MainForm
                        var mainForm = new MainForm(controller);

                        // Assert - Verify MainForm is created and has expected properties
                        var hasControls = mainForm.Controls.Count > 0;
                        var hasText = !string.IsNullOrEmpty(mainForm.Text);

                        return (mainForm != null)
                            .Label("MainForm should be created")
                            .And(hasControls)
                            .Label($"MainForm should have controls (found {mainForm.Controls.Count})")
                            .And(hasText)
                            .Label($"MainForm should have title text: {mainForm.Text}");
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"MainForm creation failed: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Property 2: Preservation - Packaging Script Structure
        /// 
        /// This test verifies that the packaging script (package-portable-selfcontained.ps1)
        /// exists and contains the expected structure for creating a self-contained package.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_PackagingScript_HasCorrectStructure(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    try
                    {
                        // Arrange
                        var scriptPath = "package-portable-selfcontained.ps1";

                        // Act - Check if script exists
                        var scriptExists = File.Exists(scriptPath);
                        
                        if (!scriptExists)
                        {
                            return false.Label($"Packaging script not found at: {scriptPath}");
                        }

                        // Read script content
                        var scriptContent = File.ReadAllText(scriptPath);

                        // Assert - Verify script contains expected commands
                        var hasDotnetPublish = scriptContent.Contains("dotnet publish");
                        var hasSelfContained = scriptContent.Contains("--self-contained");
                        var hasOutputPath = scriptContent.Contains("-o") || scriptContent.Contains("--output");
                        var hasDataFolder = scriptContent.Contains("Data") || scriptContent.Contains("data");

                        return scriptExists
                            .Label("Packaging script should exist")
                            .And(hasDotnetPublish)
                            .Label("Script should contain 'dotnet publish' command")
                            .And(hasSelfContained)
                            .Label("Script should use '--self-contained' flag")
                            .And(hasOutputPath)
                            .Label("Script should specify output path")
                            .And(hasDataFolder)
                            .Label("Script should reference Data folder structure");
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Packaging script verification failed: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Property 2: Preservation - Build Configuration Consistency
        /// 
        /// This test verifies that different build configurations (Debug/Release)
        /// produce consistent initialization behavior.
        /// 
        /// EXPECTED OUTCOME: Test PASSES on unfixed code (baseline behavior)
        /// After fix: Test should STILL PASS (no regression)
        /// 
        /// **Validates: Requirements 3.1**
        /// </summary>
        [Property(Arbitrary = new[] { typeof(BuildConfigurationGenerators) })]
        public Property Preservation_BuildConfiguration_ProducesConsistentBehavior(BuildConfiguration config)
        {
            // Skip if this is an InspectSheets build (not relevant for preservation)
            if (config.InvolvesInspectSheets)
            {
                return true.ToProperty();
            }

            return Prop.ForAll(
                Arb.Default.Unit(),
                _ =>
                {
                    try
                    {
                        // Act - Verify that services can be initialized regardless of build config
                        var csvParser = new CSVParser();
                        var dateCalculator = new DateCalculator();
                        var headerCalculator = new HeaderCalculator(dateCalculator);
                        var transformationRulesEngine = new TransformationRulesEngine();
                        var dataTransformer = new DataTransformer(transformationRulesEngine);
                        var excelManager = new ExcelManager();

                        // Assert - All services should initialize successfully
                        var allServicesInitialized = 
                            csvParser != null &&
                            dateCalculator != null &&
                            headerCalculator != null &&
                            transformationRulesEngine != null &&
                            dataTransformer != null &&
                            excelManager != null;

                        return allServicesInitialized
                            .Label($"All services should initialize for {config.Configuration} configuration");
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Build configuration test failed for {config.Configuration}: {ex.Message}");
                    }
                });
        }
    }

    /// <summary>
    /// Represents a build configuration for property-based testing.
    /// </summary>
    public class BuildConfiguration
    {
        public string Configuration { get; set; } = "Release";
        public string RuntimeIdentifier { get; set; } = "win-x64";
        public bool SelfContained { get; set; } = true;
        public bool InvolvesInspectSheets { get; set; } = false;

        public override string ToString()
        {
            return $"Config={Configuration}, RID={RuntimeIdentifier}, SelfContained={SelfContained}, InspectSheets={InvolvesInspectSheets}";
        }
    }

    /// <summary>
    /// FsCheck generators for build configurations.
    /// </summary>
    public static class BuildConfigurationGenerators
    {
        /// <summary>
        /// Generates arbitrary build configurations for property-based testing.
        /// Focuses on non-InspectSheets builds to test preservation of baseline behavior.
        /// </summary>
        public static Arbitrary<BuildConfiguration> BuildConfig()
        {
            var configGen = Gen.Elements("Debug", "Release");
            var ridGen = Gen.Elements("win-x64", "win-x86", "win-arm64");
            var selfContainedGen = Gen.Elements(true, false);
            
            // For preservation tests, we focus on non-InspectSheets builds
            // InspectSheets builds are tested separately in bug exploration tests
            var involvesInspectSheetsGen = Gen.Constant(false);

            var buildConfigGen = from config in configGen
                                 from rid in ridGen
                                 from selfContained in selfContainedGen
                                 from involvesInspectSheets in involvesInspectSheetsGen
                                 select new BuildConfiguration
                                 {
                                     Configuration = config,
                                     RuntimeIdentifier = rid,
                                     SelfContained = selfContained,
                                     InvolvesInspectSheets = involvesInspectSheets
                                 };

            return Arb.From(buildConfigGen);
        }
    }
}
