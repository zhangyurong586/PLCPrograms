// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     UnitTestScript.cs
// FileType:     Visual C# Source file
// Author:       Andre Yost, Rockwell Automation Engineering
// Created:      2024
// Description:  This script conducts unit testing for AOIs utilizing Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//               
// The main program of this script takes 3 inputs:
//   Input 1. The file path to the input excel sheet that defines the test cases & how the unit test is conducted.
//   Input 2. The file path to the output excel sheet that contains the test results.
//   Input 3. A boolean value ('true' or 'false') that determines whether to execute the slow or fast test.
//            - Fast Test: set all inputs at the same time (benefit: save time for complex tags with nested subcomponents)
//            - Slow Test: set all inputs one at a time (benefit: catch controller faults per parameter instead of per each test case)
//
// Example 1:
// .\AOIUnitTesting_ConsoleApp.exe "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\1_InputExcelFiles\WetBulbTemperature_ControllerFaultCase.xlsx"
//            In this example, only 1 input was specified so the default values for inputs 2 & 3 are used.
//            Default input 2: The output excel file is created at the input excel file's parent directory, within a new 'X_UnitTestResults' folder.
//            Default input 2 for this example: "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\X_UnitTestResults\20240816171211_UnitTestReport.xlsx"
//            Default input 3: Run the fast test ('true').
//
// Example 2:
// .\AOIUnitTesting_ConsoleApp.exe "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\1_InputExcelFiles\WetBulbTemperature_ControllerFaultCase.xlsx"
// "C:\Users\ASYost\Desktop\GeneratedTestResults"
//            In this example, only 2 inputs are specified so a default value for input 3 is used.
//            Note for input 2: If the output excel file exists at the file path provided, add the report of test results within a new worksheet.
//                              If the output excel file does not exist at the file path provided, create a new workbook and add the test results worksheet.
//            Default input 3: Run the fast test ('true').
//
// Example 3:
// .\AOIUnitTesting_ConsoleApp.exe "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\1_InputExcelFiles\WetBulbTemperature_ControllerFaultCase.xlsx"
// "C:\Users\ASYost\Desktop\GeneratedTestResults" "false"
//            In this example, all 3 inputs are specified. This will execute like example two, but with the slow test.
//            Note for input 2: If the output excel file exists at the file path provided, add the report of test results within a new worksheet.
//                              If the output excel file does not exist at the file path provided, create a new workbook and add the test results worksheet.
//            Input 3: Run the slow test ('false').
// 
// Example 4:
// .\AOIUnitTesting_ConsoleApp.exe "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\1_InputExcelFiles\WetBulbTemperature_ControllerFaultCase.xlsx"
// "" "false"
//            In this example, all 3 inputs are specified but because input 2 is an empty string, the output file path default value is used.
//            Default input 2: The output excel file is created at the input excel file's parent directory, within a new 'X_UnitTestResults' folder.
//            Default input 2 for this example: "C:\Users\ASYost\Desktop\20240816_AOIUnitTestProgress\X_UnitTestResults\20240816171211_UnitTestReport.xlsx"
//            Input 3: Run the slow test ('false').
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using Google.Protobuf;
using L5Xfiles;
using LogixEcho;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using RockwellAutomation.LogixDesigner.Logging;
using System.Collections;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static RockwellAutomation.LogixDesigner.LogixProject;

namespace UnitTest
{
    /// <summary>
    /// This class contains the methods and logic to programmatically conduct unit testing for Studio 5000 Logix Designer Add-On Instructions (AOIs).
    /// </summary>
    public class LogixUnitTestMethods
    {
        #region STRUCTURES & STATIC VARIABLES

        /// <summary>
        /// The "AOI Parameter" structure houses all the information required to read & use a single parameter of an AOI.<br/>
        /// (Note that this structure will always be used in a list, wherein each element pertains to an AOI parameter.)
        /// </summary>
        public struct AOIParameter
        {
            public string? Name { get; set; }       // The AOI parameter's name.
            public string? DataType { get; set; }   // Currently supported data types: BOOL/SINT/INT/DINT/LINT/REAL
            public string? Usage { get; set; }      // The 3 parameter usage types: Input/Output/InOut
            public bool? Required { get; set; }     // Is the parameter required in an instruction? (true/false)
            public bool? Visible { get; set; }      // Is the parameter visible in an instruction? (true/false)
            public string? Value { get; set; }      // The AOI parameter's value.
            public int BytePosition { get; set; }   // Used to track the position of the parameter in the incoming AOI byte string.
            public int BoolPosition { get; set; }   // Used to track the position of booleans in the incoming AOI byte string.
            public string? XPath { get; set; }      // The Studio 5000 tag's XPath.

            public AOIParameter() // Set default values.
            {
                Name = "";
                DataType = "";
                Usage = "";
                Required = false;
                Visible = false;
                Value = "";
                BytePosition = 0;
                BoolPosition = 0;
                XPath = "";
            }
        }

        /// <summary>
        /// The "Studio 5000 Logix Designer Tag" structure houses all the information required to read & use a single tag.
        /// </summary>
        public struct S5kAtomicTag
        {
            public string? Name { get; set; }       // The Studio 5000 tag's name.
            public string? DataType { get; set; }   // Currently supported data types: BOOL/SINT/INT/DINT/LINT/REAL
            public string? Usage { get; set; }      // Is this tag an Input or Output for unit testing
            public string? OnlineVal { get; set; }  // The Studio 5000 tag's online value.
            public string? OfflineVal { get; set; } // The Studio 5000 tag's offline value.
            public string? XPath { get; set; }      // The Studio 5000 tag's XPath.

            public S5kAtomicTag() // Set default values.
            {
                Name = "";
                DataType = "";
                Usage = "";
                OnlineVal = "";
                OfflineVal = "";
                XPath = "";
            }
        }

        // "STATIC VARIABLES" - Use to configure unit test "back-end" setup as desired.
        public static readonly int consoleCharLengthLimit = 110; // -------------------------- The character length limit of each line printed to the console.
        public static readonly string chassisName = "UnitTest_Chassis"; // ------------------- Emulated chassis name.
        public static readonly string controllerName = "UnitTest_Controller"; // ------------- Emulated controller name.
        public static readonly string processorType = "1756-L85E"; // ------------------------ The type of emulated controller used to host test.
        public static readonly string taskName = "T00_UnitTesting"; // ----------------------- Name of the continuous task in the Studio 5000 application.
        public static readonly string programName = "P00_UnitTesting"; // -------------------- Name of the program in the Studio 5000 application.
        public static readonly string routine0Name = "R00_MainRoutine"; // ------------------- Name of the main routine in the Studio 5000 application.
        public static readonly string routine1Name = "R01_UnitTesting"; /* ------------------- Name of the routine generated from AOI definition & imported to
                                                                                               the Studio 5000 application. */
        public static readonly string programName_FaultHandler = "PXX_FaultHandler"; // ------ Name of the fault handling program in the Studio 5k application.
        public static readonly string routineName_FaultHandler = "RXX_FaultHandler"; // ------ Name of the fault handling routine in the Studio 5k application.
        public static readonly DateTime testStartTime = DateTime.Now; /* --------------------- The time during which this test was first initiated. 
                                                                                               (Used at end of test to calculate unit test length.) */
        public static readonly string currentDateTime = testStartTime.ToString("yyyyMMddHHmmss"); /* Time during which test was first initiated, as a string.
                                                                                                     (Used to name generated files & test reports.)*/
        public static readonly bool conversionPrintOut = false; /* --------------------------- Making this true will print to console each step
                                                                                               taken when converting the AOI L5X to a rung L5X. */
        public static readonly bool showFullEventLog = false; /* ----------------------------- Capture and print event logger information to the console.
                                                                                               (Useful during troubleshooting.) */
        public static readonly bool deleteEchoChassis = true; /* ----------------------------- Choose whether to keep or delete emulated chassis (including 
                                                                                               its controllers) at the end of testing.*/
        #endregion

        /// <summary>
        /// This unit test example has the following steps.<br/>
        /// 1. The "input excel sheet" is parsed. This excel sheet contains the following information:<br/>
        ///    -  The Studio 5000 component is being unit tested (options: AOI_Definition.L5X, Rung.L5X, Program.L5X, Application.ACD).<br/>
        ///    -  The file path of the Studio 5000 component being tested.<br/>
        ///    -  A boolean value whether or not to retain generated ACD files.<br/>
        ///    -  A boolean value whether or not to retain generated L5X files.<br/>
        ///    -  Test cases specifying what inputs to change and what outputs to test (1 test case per excel column).<br/>
        ///    -  The number of controller clock cycles to progress each test case before verifying the outputs.<br/>
        /// 2. Create an emulated controller and chassis using the Echo SDK if one doesn't already exist.<br/>
        /// 3. A Studio 5000 Logix Designer ACD application file is created to host unit testing for L5X test inputs.<br/>
        ///    (Note: If testing ACD application, skip this section.)<br/>
        ///    -  An L5X file containing a fault handler program (contents stored within this c# solution) is converted into an ACD file.<br/>
        ///    -  If testing an AOI definition, the AOI's definition L5X is programmatically converted into a Studio 5000 rung containing a<br/>
        ///    populated instance of the AOI instruction (all required/visible instruction inputs are populated). It is then import to the ACD file.<br/>
        ///    -  If testing a rung/program L5X, import the L5X component to the ACD file.<br/>
        /// 4. Commence testing. While online with the emulated controller, the LDSDK is used to change the input parameters/tags,<br/>
        ///    then verify expected vs. actual output parameter results.<br/>
        /// 5. Put unit test results into a worksheet of an excel workbook.<br/>
        ///    If the excel workbook specified in the input command does not yet exist, the workbook is created.<br/>
        ///    If the excel workbook specified in the input command exists, a new worksheet is added to the workbook.<br/>
        ///    (Note for potential future modifications of this unit test script: the output excel sheet containing the results of the<br/>
        ///     unit test was programmatically created and modified at 4 separate locations of this script.)        
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static async Task Main(string[] args)
        {
            #region PARSE VARIABLES & INITIALIZE UNIT TEST
            bool issueWithScriptInputParsing = false; // A boolean value used to stop script if an issue parsing inputs is encountered.

            // REQUIRED VALUE: input string 1
            // The input excel workbook file path. (This file defines the test cases & how the unit test is conducted).
            string inputArg_inputExcelFilePath = args[0];

            // DEFAULT VALUE: input string 2
            // If no output excel folder path is provided, use the below file path at which to create the test report.
            // If an output excel folder path is provided, overwrite the below value.
            string outputExcelFolderPath = Directory.GetParent(Path.GetDirectoryName(inputArg_inputExcelFilePath)!) + @"\X_UnitTestResults\";
            if ((args.Length == 1) && (!Directory.Exists(outputExcelFolderPath)))
            {
                Directory.CreateDirectory(outputExcelFolderPath); // If default output, folder 'UnitTestResults' created at input excel file's parent directory
            }
            string inputArg_outputExcelFilePath = outputExcelFolderPath + currentDateTime + "_UnitTestReport.xlsx";

            // OVERRIDE DEFAULT VALUE: input string 2
            // The output excel workbook file path. (This file will contain the results of unit testing).
            // Note: If file does not exist, create it. If it exists, add a new worksheet to the existing workbook.
            if ((args.Length > 1) && (args[1] != "")) inputArg_outputExcelFilePath = args[1];

            // DEFAULT VALUE: input string 3
            // If no boolean value 'true' or 'false' (capitalization does not matter) is provided, set default unit test functionality to a fast test.
            bool inputArg_fastTest = true;

            // Handle any potential issues with console application input scenarios.
            // Scenario 1: Incorrect # of inputs.
            // Scenario 2: Input excel file does not exist.
            // Scenario 3: 3rd argument is not formatted properly as a boolean input.
            if ((args.Length < 1) || (args.Length > 3))     // Scenario 1
            {
                ConsoleMessage("INCORRECT NUMBER OF INPUTS", "ERROR");
                Console.Write(@"Correct Command Example: .\LogixUnitTesting inputExcelWorkbook_FilePath 
                                outputExcelWorkbook_FilePath booleanValueForFastOrSlowTest");
                issueWithScriptInputParsing = true;
            }
            if (!File.Exists(inputArg_inputExcelFilePath))  // Scenario 2
            {
                ConsoleMessage("Input excel workbook directory does not exist.", "ERROR");
                issueWithScriptInputParsing = true;
            }

            // OVERRIDE DEFAULT VALUE: input string 3
            // The value determining the slow or fast execution of unit testing.
            // Fast Test: set all inputs at the same time (benefit: save time for AOIs/complex tags with many parameters)
            // Slow Test: set all inputs one at a time (benefit: catch controller faults per parameter/tag instead of per each test case)
            if (args.Length > 2)
            {
                if (bool.TryParse(args[2], out bool result))
                    inputArg_fastTest = result;
                else                                        // Scenario 3
                {
                    ConsoleMessage("Invalid boolean value for the third argument. Use either 'true' or 'false' (uppercase/lowercase does not matter).",
                        "ERROR");
                    issueWithScriptInputParsing = true;
                }
            }

            // End the main method script if an issue was encountered during input variable parsing.
            if (issueWithScriptInputParsing)
                return;

            // Create the folder that will contain the generated ACD files and/or L5X files. Note that the folder is deleted, if empty, at the end of the test.
            string generatedFilesFolderPath = Directory.GetParent(Path.GetDirectoryName(inputArg_outputExcelFilePath)!) + @"\X_GeneratedFiles\";
            if (!Directory.Exists(generatedFilesFolderPath))
                Directory.CreateDirectory(generatedFilesFolderPath);

            // Print unit test banner to the console.
            Console.WriteLine("\n  ".PadRight(consoleCharLengthLimit - 2, '='));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            string bannerContents = "UNIT TESTING | " + DateTime.Now + " " + TimeZoneInfo.Local;
            int padding = (consoleCharLengthLimit - bannerContents.Length) / 2;
            Console.WriteLine(bannerContents.PadLeft(bannerContents.Length + padding).PadRight(consoleCharLengthLimit));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            Console.WriteLine("  ".PadRight(consoleCharLengthLimit - 2, '=') + "\n");

            // Print the input argument test parameters to the console.
            ConsoleMessage("START parsing input arguments for unit testing...", "NEWSECTION", false);
            ConsoleMessage($"Input excel sheet file path to be used: '{inputArg_inputExcelFilePath}'", "STATUS");
            ConsoleMessage($"Output excel sheet file path to be used: '{inputArg_outputExcelFilePath}'", "STATUS");
            ConsoleMessage(inputArg_fastTest ? "Fast test selected." : "Slow test selected.", "STATUS");
            ConsoleMessage("The fast test sets all inputs at the same time (benefit: save time for complex tags with nested subcomponents).");
            ConsoleMessage("The slow test sets all inputs one at a time (benefit: catch controller faults per parameter instead of per each test case).");

            ConsoleMessage("START parsing input excel workbook test information, parameters, and test cases...", "NEWSECTION");

            // Variables containing information about the object file to test & about whether to retain generated ACD or L5X files.
            string iExcel_testObjectType = "AOI_DEFINITION.L5X";
            string iExcel_testObjectFilePath;
            bool iExcel_keepACDs;
            bool iExcel_keepL5Xs;

            // Populate the above variables from the input excel file.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_inputExcelFilePath)))
            {
                ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                iExcel_testObjectFilePath = inputExcelWorksheet.Cells[9, 2].Value.ToString()!.Trim()!;
                iExcel_keepACDs = bool.Parse(inputExcelWorksheet.Cells[9, 16].Value?.ToString()!.Trim()!);
                iExcel_keepL5Xs = bool.Parse(inputExcelWorksheet.Cells[9, 21].Value?.ToString()!.Trim()!);
            }

            // Print message to console about what kind of unit test is being conducted.
            if (iExcel_testObjectType == "APPLICATION.ACD")
                ConsoleMessage("Full ACD application unit test selected.", "STATUS");
            else if (iExcel_testObjectType == "AOI_DEFINITION.L5X")
                ConsoleMessage("AOI definition L5X unit test selected.", "STATUS");
            else if (iExcel_testObjectType == "RUNG.L5X")
                ConsoleMessage("Rung unit test selected.", "STATUS");
            else if (iExcel_testObjectType == "ROUTINE.L5X")
                ConsoleMessage("Routine L5X unit test selected.", "STATUS");
            else
            {
                ConsoleMessage($"Test object type '{iExcel_testObjectType}' not supported. Select either AOI_Definition.L5X, Rung.L5X, Program.L5X, or " +
                    $"Application.ACD.", "ERROR");
                return;
            }

            // Print message to console about the selected input excel test information.
            ConsoleMessage($"File to be tested: '{iExcel_testObjectFilePath}'.", "STATUS", false);
            if (iExcel_testObjectType != "APPLICATION.ACD")
            {
                if (iExcel_keepACDs)
                    ConsoleMessage($"Retain generated ACD files used to host unit tests.", "STATUS", false);
                else
                    ConsoleMessage($"Delete ACD files used to host unit tests.", "STATUS", false);

                if (iExcel_keepL5Xs)
                    ConsoleMessage($"Retain generated L5X files used to set up unit test.", "STATUS", false);
                else
                    ConsoleMessage($"Delete generated L5X files used to set up unit test.", "STATUS", false);
            }
            #endregion

            #region STAGING TEST: create new ACD -> create emulated controller & chassis -> import L5Xs to ACD -> download ACD -> put controller in run mode
            string commPath = "";
            string partialL5XprojectFilePath = "";
            string newAOIroutineL5XFilePath = "";
            string acdFilePath = "";
            string testObjectName = "";
            LogixProject logixProject;

            if (iExcel_testObjectType != "APPLICATION.ACD")
            {
                ConsoleMessage("START creating & opening ACD application file to be used during testing...", "NEWSECTION");

                // Get variables needed to set up unit test. Information retreived from the L5X or ACD file specified in the input excel sheet.
                if (iExcel_testObjectType == "AOI_DEFINITION.L5X")
                    testObjectName = GetAttributeValue(iExcel_testObjectFilePath, "AddOnInstructionDefinition", "Name", false)!;
                else
                    testObjectName = GetAttributeValue(iExcel_testObjectFilePath, "RSLogix5000Content", "TargetName", false)!;
                string softwareRevision = GetAttributeValue(iExcel_testObjectFilePath, "RSLogix5000Content", "SoftwareRevision", false)!;

                /* Create the ACD file to host unit test.
                 * (Note that these steps are necessary in order to include a program within 'Controller Fault Handler' in Studio 5k.)
                     Step 1. Get a string containing the L5X application contents needed to make an L5X file.
                     Step 2. Create a newly generated L5X application file.
                     Step 3. Open L5X file using LDSDK (has to be opened for step 4).
                     Step 4. Convert the open L5X file to ACD file using the LDSDK. 
                     Step 5. Open ACD file using LDSDK.*/
                string l5xFileContents = L5XFiles.GetFaultHandlingApplicationL5XContents(routine0Name, routine1Name, programName, taskName,
                    routineName_FaultHandler, programName_FaultHandler, controllerName, processorType, softwareRevision);    // Step 1: Get L5X contents.
                partialL5XprojectFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_Basic.L5X";
                File.WriteAllText(partialL5XprojectFilePath, l5xFileContents);                                                             // Step 2: Generate new L5X file.
                LogixProject logixProjectL5X = await LogixProject.OpenLogixProjectAsync(partialL5XprojectFilePath);                        // Step 3: Open L5X file.
                ConsoleMessage($"L5X application file created & opened at '{partialL5XprojectFilePath}'.", "STATUS");
                acdFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_WithAOI.ACD";
                await logixProjectL5X.SaveAsAsync(acdFilePath, true);                                                        // Step 4: Convert L5X to ACD.
                logixProject = await LogixProject.OpenLogixProjectAsync(acdFilePath);                           // Step 5: Open ACD file.
                ConsoleMessage($"ACD application file created & opened at '{acdFilePath}'.", "STATUS");

                // Capture and print event logger information to the console. (Useful during troubleshooting.)
                if (showFullEventLog)
                    logixProject.AddEventHandler(new StdOutEventLogger());

                // Set up emulated controller (based on the specified ACD file path & unit test static variables) if one does not yet exist.
                ConsoleMessage("START setting up Factory Talk Logix Echo emulated controller...", "NEWSECTION");
                commPath = LogixEchoMethods.Main(acdFilePath, chassisName, controllerName).GetAwaiter().GetResult();
                ConsoleMessage($"Project communication path specified is '{commPath}'.", "STATUS");

                // Finish setting up ACD application for unit testing by importing the AOI & AOI rung L5X files.
                ConsoleMessage("START preparing ACD application environment for unit test...", "NEWSECTION");
                string xPath_aoiDef = @"Controller/AddOnInstructionDefinitions";
                await logixProject.PartialImportFromXmlFileAsync(xPath_aoiDef, iExcel_testObjectFilePath,                 // Import the AOI.L5X being tested
                    LogixProject.ImportCollisionOptions.OverwriteOnColl);                                                    // to the open ACD application.
                await logixProject.SaveAsync();
                ConsoleMessage($"Imported '{iExcel_testObjectFilePath}' to '{acdFilePath}'.", "STATUS");

                if (iExcel_testObjectType == "AOI_DEFINITION.L5X")
                {
                    // Convert a copy of the AOI.L5X into rung.L5X format, then import into the ACD application. 
                    // The ladder logic rung contains an instance of the AOI instruction populated with any visible and/ or required tags.
                    ConsoleMessage($"Print STATUS messages for AOI.L5X to rung.L5X conversion? Currently set to '{conversionPrintOut}'.", "STATUS");
                    newAOIroutineL5XFilePath = CopyFile(iExcel_testObjectFilePath, generatedFilesFolderPath);
                    ConvertL5X_AOItoROUTINE(newAOIroutineL5XFilePath, routine1Name, programName, controllerName, conversionPrintOut); // Convert AOIDefinition.L5X to routine.L5X
                }

                string xPath_convertedRungFromAOI = @"Controller/Programs";
                await logixProject.PartialImportFromXmlFileAsync(xPath_convertedRungFromAOI, newAOIroutineL5XFilePath,            // Import the programmatically created
                    LogixProject.ImportCollisionOptions.OverwriteOnColl);                                               // rung to the open ACD application.
                await logixProject.SaveAsync();
                ConsoleMessage($"Imported '{newAOIroutineL5XFilePath}' to '{acdFilePath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage("START opening ACD application file to be used during testing...", "NEWSECTION");
                logixProject = await LogixProject.OpenLogixProjectAsync(iExcel_testObjectFilePath);
            }


            // Change emulated controller mode to program & verify.
            ConsoleMessage("START changing controller to PROGRAM...", "NEWSECTION");
            ChangeControllerMode_Async(commPath, "PROGRAM", logixProject).GetAwaiter().GetResult();
            if (ReadControllerMode_Async(commPath, logixProject).GetAwaiter().GetResult() == "PROGRAM")
                ConsoleMessage("SUCCESS changing controller to PROGRAM.", "STATUS", false);
            else
                ConsoleMessage("FAILURE changing controller to PROGRAM.", "ERROR", false);

            // Download programmatically generated ACD application to the emulated controller.
            ConsoleMessage("START downloading ACD file...", "NEWSECTION");
            DownloadProject_Async(commPath, logixProject).GetAwaiter().GetResult();
            ConsoleMessage("SUCCESS downloading ACD file.", "STATUS", false);

            // Change emulated controller mode to run & verify.
            ConsoleMessage("START changing controller to RUN...", "NEWSECTION");
            ChangeControllerMode_Async(commPath, "RUN", logixProject).GetAwaiter().GetResult();
            if (ReadControllerMode_Async(commPath, logixProject).GetAwaiter().GetResult() == "RUN")
                ConsoleMessage("SUCCESS changing controller to RUN.", "STATUS", false);
            else
                ConsoleMessage("FAILURE changing controller to RUN.", "ERROR", false);

            string fullL5XprojectFilePath = generatedFilesFolderPath + currentDateTime + "_" + testObjectName + "_FullProj.L5X";
            await logixProject.SaveAsAsync(fullL5XprojectFilePath, true);
            #endregion

            #region COMMENCE TEST: Set & check parameters for each test case from the excel sheet. Results are committed to output excel worksheet.
            ConsoleMessage($"START {testObjectName} unit testing...", "NEWSECTION");

            // Get the Name, DataType, Usage, Required, and Visible components of each parameter from the AOI definition XML file
            // & put those contents into an array.
            AOIParameter[] AOIParameters = GetComplexParameters_FromL5X(iExcel_testObjectFilePath, testObjectName)!;

            // Store the XPath of the AOI Studio 5000 Logix Designer tag that was programmatically created and used during testing.
            string aoiTagXPath = $"Controller/Tags/Tag[@Name='AOI_{testObjectName}']";

            // Unit test variables
            S5kAtomicTag AT_FaultType;         // AT_FaultType tag storing the controller fault type information.
            S5kAtomicTag AT_FaultCode;         // AT_FaultCode tag storing the controller fault code information.
            bool faultedState = false;         // An indicator of whether the controller is faulted or not.
            bool breakOutputParameterLoop;     // Used to break the "OUTPUT PARAMETER LOOP" if controller faulted when setting inputs. 
            int testCases = GetNumberOfTestCases(inputArg_inputExcelFilePath, 19, 6); // The number of test cases provided in the excel input workbook.
            int failureCondition = 0;          // This variable tracking the number of failed test cases or controller faults.
            string previousTestEnableIn = "0"; /* Track the previous value of the EnableIn parameter.
                                                  Used in logic determining whether or not the tag AT_EnableIn needs to be updated. */

            #region OUTPUT EXCEL REPORT (location 1/4 where workbook is updated): setting up & formatting output excel with banners & row names
            int lowerColumnLimit = 13 + AOIParameters.Length * 3; // Used for excel sheet formatting.

            if (iExcel_testObjectType == "AOI_DEFINITION.L5X")
            {


                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_outputExcelFilePath)))
                {
                    // Store all test information in a worksheet with a uniquely time-stamped name.
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{currentDateTime}_{testObjectName}");

                    worksheet.Cells["B2:O6"].Merge = true;
                    worksheet.Cells["B2:O6"].Value = $"{testObjectName} AOI Unit Test Results";
                    worksheet.Cells["B2:O6"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                    worksheet.Cells["B2:O6"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells["B2:O6"].Style.Font.Size = 26;
                    worksheet.Cells["B2:O6"].Style.Font.Bold = true;
                    worksheet.Cells["B2:O6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B2:O6"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["B8"].Value = "Overall Test Result:";
                    worksheet.Cells["B8"].Style.Font.Bold = true;
                    worksheet.Cells["B8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    worksheet.Cells["C8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    worksheet.Cells["E8:F8"].Merge = true;
                    worksheet.Cells["E8:F8"].Value = "Date Test Run:";
                    worksheet.Cells["E8:F8"].Style.Font.Bold = true;
                    worksheet.Cells["E8:F8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    worksheet.Cells["G8:I8"].Merge = true;
                    worksheet.Cells["G8:I8"].Value = testStartTime.ToString();
                    worksheet.Cells["G8:I8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    worksheet.Cells["K8:M8"].Merge = true;
                    worksheet.Cells["K8:M8"].Value = "Total Test Time (hh:mm:ss):";
                    worksheet.Cells["K8:M8"].Style.Font.Bold = true;
                    worksheet.Cells["K8:M8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    worksheet.Cells["B10"].Value = "Test Cases:";
                    worksheet.Row(10).Style.Font.Bold = true;
                    worksheet.Cells["B10"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    for (int i = 0; i < testCases; i++)
                    {
                        worksheet.Cells[10, 3 + i].Value = i + 1;
                    }
                    worksheet.Cells[10, 3, 10, 3 + testCases].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells[11, 3, 11, 3 + testCases].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B11"].Value = "Test Case Result:";
                    worksheet.Row(11).Style.Font.Bold = true;
                    worksheet.Cells["B11"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    worksheet.Cells["B13"].Value = "Tested Input Parameters:";
                    worksheet.Cells["B13"].Style.Font.Bold = true;

                    worksheet.Cells[13, 2, lowerColumnLimit, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    int rowNum = 13;
                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Output")
                        {
                            rowNum++;
                            worksheet.Cells[rowNum, 2].Value = parameter.Name;
                        }
                    }

                    rowNum = rowNum + 2;
                    worksheet.Cells[rowNum, 2].Value = "Tested Output Parameters:";
                    worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Input")
                        {
                            rowNum++;
                            worksheet.Cells[rowNum, 2].Value = parameter.Name;
                        }
                    }

                    rowNum = rowNum + 2;
                    worksheet.Cells[rowNum, 2].Value = "Actual Output Parameters:";
                    worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Input")
                        {
                            rowNum++;
                            worksheet.Cells[rowNum, 2].Value = parameter.Name;
                        }
                    }

                    rowNum = rowNum + 2;
                    worksheet.Cells[rowNum, 2].Value = "Controller Fault Info:";
                    worksheet.Cells[rowNum, 2].Style.Font.Bold = true;
                    rowNum++;
                    worksheet.Cells[rowNum, 2].Value = "Type";
                    rowNum++;
                    worksheet.Cells[rowNum, 2].Value = "Code";

                    worksheet.Column(2).AutoFit();

                    package.Save();
                }
            }
            ConsoleMessage($"Done setting up & formatting output excel test report.", "STATUS");
            #endregion

            // Get the values of the "Safe State" test case #0 for each AOI parameter.
            Dictionary<string, string> safeStateTestCase = GetExcelTestCaseValues(inputArg_inputExcelFilePath, 5);

            // Set values to known safe state
            await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);
            await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);
            ConsoleMessage($"{testObjectName} parameter values set to \"Safe State\" test case #0, shown below:", "STATUS");
            PrintAOIParameters(AOIParameters, false);

            // TEST CASES LOOP: Iterate through each test case provided in the input excel workbook (each column).
            for (int i = 0; i < testCases; i++)
            {
                // Parameters updated/cleared each test case.
                int testNumber = i + 1;                  // The test case currently being tested.
                int inputExcelColumnNum = i + 6;        // The number of the input excel column from which test case values are being obtained.
                int faultType = 0;                       // Integer variable storing the controller fault type number (used in output excel).
                int faultCode = 0;                       // Integer variable storing the controller fault code number (used in output excel).
                breakOutputParameterLoop = false;        // Used to break the "OUTPUT PARAMETER LOOP" if controller faulted when setting inputs.
                int currentColumnNumForOutExcel = i + 3; // Required value for programmatically creating output excel file.
                int testIfFailure = failureCondition;    // testIfFailure used as a comparison value for whether an individual test case failed.

                ConsoleMessage($"START test case {testNumber}/{testCases}...", "NEWSECTION", false);

                // Set values to the known safe state, test case #0 of the input excel sheet.
                await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);
                ConsoleMessage($"Parameters set to \"Safe State\" test case #0.", "STATUS");

                // The EnableIn parameter is modified by changing the boolean tag AT_EnableIn (within an XIC instruction before the AOI instruction).
                if ((previousTestEnableIn == "0") && (safeStateTestCase["EnableIn"] == "1"))
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "true", OperationMode.Online, DataType.BOOL, logixProject);
                else if ((previousTestEnableIn == "1") && (safeStateTestCase["EnableIn"] == "0"))
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "false", OperationMode.Online, DataType.BOOL, logixProject);

                // Forward the test by 1 controller clock cycle.
                // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                // The JSR runs the programmatically generated routine containing the AOI instruction.)
                await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);
                ConsoleMessage($"Progressed controller by 1 clock cycle.", "STATUS");

                // Get the current test case values to be used during testing.
                Dictionary<string, string> currentTestCaseValues = GetExcelTestCaseValues(inputArg_inputExcelFilePath, inputExcelColumnNum);

                // The EnableIn parameter is modified by changing the boolean tag AT_EnableIn (within an XIC instruction before the AOI instruction).
                if ((safeStateTestCase["EnableIn"] == "0") && (currentTestCaseValues["EnableIn"] == "1"))
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "true", OperationMode.Online, DataType.BOOL, logixProject);
                else if ((safeStateTestCase["EnableIn"] == "1") && (currentTestCaseValues["EnableIn"] == "0"))
                    await SetTagValue_Async("Controller/Tags/Tag[@Name='AT_EnableIn']", "false", OperationMode.Online, DataType.BOOL, logixProject);

                #region UNIT TEST: change all AOI input parameters for the given test case (execute FAST or SLOW given the specified input argument for script)
                ConsoleMessage($"Setting input parameter values for test case {testNumber}/{testCases}.", "STATUS");

                // Fast Test: set all inputs at the same time (benefit: save time for complex tags with nested subcomponents)
                if (inputArg_fastTest)
                {
                    // Set all input parameters based on the current test case values.
                    await SetMultipleAOIParamVals_Async(aoiTagXPath, currentTestCaseValues, AOIParameters, OperationMode.Online, logixProject, true);

                    // Forward the test by 1 controller clock cycle.
                    // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                    // The JSR runs the programmatically generated routine containing the AOI instruction.)
                    ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject).GetAwaiter().GetResult();
                    ConsoleMessage($"Progressed controller by 1 clock cycle.", "STATUS");

                    // Check if changing the input parameters for this test case caused a controller fault.
                    AT_FaultType = GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultType']", DataType.DINT, logixProject).GetAwaiter().GetResult();
                    AT_FaultCode = GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultCode']", DataType.DINT, logixProject).GetAwaiter().GetResult();
                    faultType = int.Parse(AT_FaultType.OnlineVal!); // Used in output excel sheet report.
                    faultCode = int.Parse(AT_FaultCode.OnlineVal!); // Used in output excel sheet report.
                    faultedState = (AT_FaultType.OnlineVal! != "0") || (AT_FaultCode.OnlineVal! != "0");

                    if (faultedState)
                        ConsoleMessage($"Controller faulted with type #{AT_FaultType.OnlineVal} & code #{AT_FaultCode.OnlineVal}.", "FAIL");
                }

                // Slow Test: set all inputs one at a time (benefit: catch controller faults per parameter instead of per each test case)
                else if (!inputArg_fastTest)
                {
                    // SLOW TEST SET INPUT PARAMETERS LOOP
                    foreach (var kvp in currentTestCaseValues)
                    {
                        if (GetAOIParameterComponentValue(kvp.Key, "Usage", AOIParameters) == "Input")
                        {
                            if (GetAOIParameterComponentValue(kvp.Key, "DataType", AOIParameters) == "BOOL")
                            {
                                await SetSingleAOIParamValue_Async(aoiTagXPath, (kvp.Value == "1").ToString(), kvp.Key, OperationMode.Online, AOIParameters,
                                    logixProject, true);
                            }
                            else
                            {
                                await SetSingleAOIParamValue_Async(aoiTagXPath, kvp.Value, kvp.Key, OperationMode.Online, AOIParameters, logixProject, true);
                            }

                            // Forward the test by 1 controller clock cycle.
                            // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                            // The JSR runs the programmatically generated routine containing the AOI instruction.)
                            await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);
                            ConsoleMessage($"Progressed controller by 1 clock cycle.", "STATUS");

                            // Check if changing the input parameters for this test case caused a controller fault.
                            AT_FaultType = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultType']", DataType.DINT, logixProject);
                            AT_FaultCode = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultCode']", DataType.DINT, logixProject);
                            faultType = int.Parse(AT_FaultType.OnlineVal!); // for excel sheet
                            faultCode = int.Parse(AT_FaultCode.OnlineVal!); // for excel sheet
                            faultedState = (AT_FaultType.OnlineVal! != "0") || (AT_FaultCode.OnlineVal! != "0");

                            if (faultedState)
                            {
                                ConsoleMessage($"Controller faulted upon setting '{kvp.Key}' to '{kvp.Value}', with fault type #{AT_FaultType.OnlineVal} & code " +
                                $"#{AT_FaultCode.OnlineVal}.", "ERROR");
                                break; // Break the "SLOW TEST SET INPUT PARAMETERS LOOP" 
                            }
                        }
                    }
                }

                // If controller faulted, attempt to clear it.
                if (faultedState)
                {
                    failureCondition++;

                    ConsoleMessage($"Attempting to clear fault. Setting all input parameter values to the \"Safe State\" test case #0 & verifying if " +
                        $"controller no longer faulted.", "STATUS");

                    // Clear the fault (toggle XIC to clear the two tags AT_FaultType & AT_FaultCode in the Studio 5000 Logix Designer Project).
                    await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ClearFault']", true, OperationMode.Online, logixProject);

                    // Set values to "Safe State" test case #0.
                    await SetMultipleAOIParamVals_Async(aoiTagXPath, safeStateTestCase, AOIParameters, OperationMode.Online, logixProject);

                    // Forward the test by 1 controller clock cycle.
                    // (AT_ToggleTest tag is in an XIC instruction followed by a ONS instruction followed by a JSR instruction.
                    // The JSR runs the programmatically generated routine containing the AOI instruction.)
                    ConsoleMessage($"Progressing controller by 1 clock cycle.", "STATUS");
                    await ToggleBOOLTagValue_Async("Controller/Tags/Tag[@Name='AT_ToggleTest']", true, OperationMode.Online, logixProject);

                    // Check if controller still faulted after setting to "Safe State" test case #0.
                    AT_FaultType = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultType']", DataType.DINT, logixProject);
                    AT_FaultCode = await GetTagValue_Async("Controller/Tags/Tag[@Name='AT_FaultCode']", DataType.DINT, logixProject);
                    faultedState = (AT_FaultType.OnlineVal! != "0") || (AT_FaultCode.OnlineVal! != "0");

                    if (faultedState)
                    {
                        ConsoleMessage("Controller still faulted. Ending Test.", "ERROR");
                        break; // Break the "TEST CASES LOOP" 
                    }
                    else if (testNumber < testCases) // Controller is not faulted and there are more tests remaining.
                    {
                        ConsoleMessage($"Fault cleared. Moving to next test case...", "SUCCESS");
                        breakOutputParameterLoop = true; // Break the "OUTPUT PARAMETER LOOP"
                    }
                }
                #endregion

                // Get the current parameter values of the Studio 5000 AOI tag with which to verify parameter outputs.
                ByteString aoiByteString = await logixProject.GetTagValueAsync(aoiTagXPath, OperationMode.Online, DataType.BYTE_ARRAY);
                AOIParameter[] AOIParameters_WithOutputs = GetAOIParameterValues(AOIParameters, aoiByteString);

                #region OUTPUT EXCEL REPORT (location 2/4 where workbook is updated): test case parameter values from input excel added to output excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_outputExcelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                    int rowNum = 13;
                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Output")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (parameter.Name == kvp.Key)
                                {
                                    rowNum++;

                                    if (parameter.DataType == "BOOL")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(kvp.Value);
                                    else if (parameter.DataType == "SINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = sbyte.Parse(kvp.Value);
                                    else if (parameter.DataType == "INT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(kvp.Value);
                                    else if (parameter.DataType == "DINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = double.Parse(kvp.Value);
                                    else if (parameter.DataType == "LINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = long.Parse(kvp.Value);
                                    else if (parameter.DataType == "REAL")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = decimal.Parse(kvp.Value);
                                    else if (parameter.DataType == "STRING")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = kvp.Value;

                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    foreach (var parameter in AOIParameters)
                    {
                        if (parameter.Usage != "Input")
                        {
                            foreach (var kvp in currentTestCaseValues)
                            {
                                if (parameter.Name == kvp.Key)
                                {
                                    rowNum++;

                                    if (parameter.DataType == "BOOL")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(kvp.Value);
                                    else if (parameter.DataType == "SINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = sbyte.Parse(kvp.Value);
                                    else if (parameter.DataType == "INT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(kvp.Value);
                                    else if (parameter.DataType == "DINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = double.Parse(kvp.Value);
                                    else if (parameter.DataType == "LINT")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = long.Parse(kvp.Value);
                                    else if (parameter.DataType == "REAL")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = decimal.Parse(kvp.Value);
                                    else if (parameter.DataType == "STRING")
                                        worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = kvp.Value;

                                    break;
                                }
                            }
                        }
                    }

                    rowNum = rowNum + 2;

                    foreach (var parameter in AOIParameters_WithOutputs)
                    {
                        if (parameter.Usage != "Input")
                        {
                            rowNum++;

                            if (parameter.DataType == "BOOL")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(parameter.Value!);
                            else if (parameter.DataType == "SINT")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = sbyte.Parse(parameter.Value!);
                            else if (parameter.DataType == "INT")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = int.Parse(parameter.Value!);
                            else if (parameter.DataType == "DINT")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = double.Parse(parameter.Value!);
                            else if (parameter.DataType == "LINT")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = long.Parse(parameter.Value!);
                            else if (parameter.DataType == "REAL")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = decimal.Parse(parameter.Value!);
                            else if (parameter.DataType == "STRING")
                                worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = parameter.Value;
                        }
                    }

                    rowNum = rowNum + 3;
                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = faultType;
                    rowNum++;
                    worksheet.Cells[rowNum, currentColumnNumForOutExcel].Value = faultCode;

                    worksheet.Column(currentColumnNumForOutExcel).AutoFit();
                    worksheet.Cells[14, currentColumnNumForOutExcel, lowerColumnLimit, currentColumnNumForOutExcel].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    package.Save();
                }
                #endregion

                // OUTPUT PARAMETER LOOP
                foreach (var kvp in currentTestCaseValues)
                {
                    if (breakOutputParameterLoop)
                        break;

                    if (GetAOIParameterComponentValue(kvp.Key, "Usage", AOIParameters) != "Input")
                    {
                        string outputValue = GetAOIParameterComponentValue(kvp.Key, "Value", AOIParameters_WithOutputs);
                        failureCondition += TEST_CompareForExpectedValue(kvp.Key, kvp.Value, outputValue, true); // If values not equal, failure condition increased.
                    }
                }

                #region OUTPUT EXCEL REPORT (location 3/4 where workbook is updated): actual output parameter values added to output excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_outputExcelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                    if (testIfFailure != failureCondition)
                    {
                        worksheet.Cells[11, currentColumnNumForOutExcel].Value = "FAIL";
                        worksheet.Cells[11, currentColumnNumForOutExcel].Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        worksheet.Cells[11, currentColumnNumForOutExcel].Value = "PASS";
                        worksheet.Cells[11, currentColumnNumForOutExcel].Style.Font.Color.SetColor(Color.Green);
                    }

                    package.Save();
                }
                ConsoleMessage($"Updated output excel test report with test case {testNumber}/{testCases}.", "STATUS");
                #endregion

                // Used in logic determining whether or not the tag AT_EnableIn needs to be updated.
                previousTestEnableIn = currentTestCaseValues["EnableIn"];
            }

            // Based on the AOI unit test result, print a final result message in red or green.
            if (failureCondition > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ConsoleMessage($"{testObjectName} Unit Test Final Result: FAIL | {failureCondition} Issues Encountered", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                ConsoleMessage($"{testObjectName} Unit Test Final Result: PASS", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            #endregion

            #region END TEST: Print final test results & retain/delete generated test components as specified in input excel sheet.
            ConsoleMessage("START retaining or deleting programmatically generated test components...", "NEWSECTION");

            // Based on the AOI Excel Worksheet for this AOI, keep or delete generated L5X files.
            if (!iExcel_keepL5Xs)
            {
                File.Delete(partialL5XprojectFilePath);
                File.Delete(newAOIroutineL5XFilePath);
                File.Delete(fullL5XprojectFilePath); // DELETE LATER: this needs to only be deleted based on the test object --------------------------------------------------------------------------------------------
                ConsoleMessage($"Deleted '{partialL5XprojectFilePath}'.", "STATUS");
                ConsoleMessage($"Deleted '{newAOIroutineL5XFilePath}'.", "STATUS");
                ConsoleMessage($"Deleted '{fullL5XprojectFilePath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{partialL5XprojectFilePath}'.", "STATUS");
                ConsoleMessage($"Retained '{newAOIroutineL5XFilePath}'.", "STATUS");
                ConsoleMessage($"Retained '{fullL5XprojectFilePath}'.", "STATUS");
            }

            // Based on the AOI Excel Worksheet for this AOI, keep or delete the generated ACD file.
            if (!iExcel_keepACDs)
            {
                File.Delete(acdFilePath);
                ConsoleMessage($"Deleted '{acdFilePath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{acdFilePath}'.", "STATUS");
            }

            // Delete all backup files (ACDs are so frequently generated that there is little value in retaining backups).
            File.Delete(acdFilePath + ".BAK");

            // Delete the generated files folder if it is empty.
            if (!Directory.EnumerateFileSystemEntries(generatedFilesFolderPath).Any())
            {
                Directory.Delete(generatedFilesFolderPath, true);
                ConsoleMessage($"Deleted empty folder '{generatedFilesFolderPath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained folder '{generatedFilesFolderPath}'.", "STATUS");
            }

            // Based on the static variable deleteEchoChassis, keep or delete the Logix Echo chassis (and its controllers) used during testing.
            if (deleteEchoChassis)
            {
                await LogixEchoMethods.DeleteChassis_Async(chassisName);
                ConsoleMessage($"Deleted Logix Echo chassis named '{chassisName}' & controller named '{controllerName}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained Logix Echo chassis named '{chassisName}' & controller named '{controllerName}'.", "STATUS");
            }

            // Compute how long the test took to run. 
            DateTime testEndTime = DateTime.Now;
            TimeSpan testLength = testEndTime.Subtract(testStartTime);
            string formattedTestLength = testLength.ToString(@"hh\:mm\:ss");
            ConsoleMessage($"AOI Unit testing for '{testObjectName}' completed in {formattedTestLength} (HH:mm:ss).", "NEWSECTION");

            #region OUTPUT EXCEL REPORT (location 4/4 where workbook is updated): test length and overall test result added
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_outputExcelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.LastOrDefault()!;

                if (failureCondition > 0)
                {
                    worksheet.Cells["C8"].Value = "FAILURE";
                    worksheet.Cells["B2:O6"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
                else
                {
                    worksheet.Cells["C8:D8"].Value = "SUCCESS";
                    worksheet.Cells["B2:O6"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                }

                worksheet.Cells["N8"].Value = formattedTestLength;
                worksheet.Cells["N8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                package.Save();
            }
            #endregion

            await logixProject.GoOfflineAsync(); // Testing is complete. Go offline with the emulated controller.

            #endregion
        }

        #region METHODS: L5X Manipulation

        /// <summary>
        /// Programmatically get the XPath of a specified tag from a Studio 5000 Logix Designer L5X file.
        /// </summary>
        /// <param name="l5xPath">The file path to a Studio 5000 Logix Designer L5X file.</param>
        /// <param name="tagName">The target tag to get the XPath of.</param>
        /// <param name="programName">
        /// If the tag is known to be a program scoped, specify the program name here.<br/>
        /// If this input is left blank and the tag name specified exists in multiple programs, the first program listed in the L5X file is used in the XPath. 
        /// </param>
        /// <returns>
        /// The XPath of a specified Studio 5000 Logix Designer tag.<br/>
        /// OR<br/>
        /// A null string if the tag name specified does not exist within the L5X file.
        /// </returns>
        private static string? GetXPathFromL5X(string l5xPath, string tagName, string? programName = null)
        {
            XDocument xDoc = XDocument.Load(l5xPath);

            // PROGRAM SCOPED XPATH SEARCH
            // Find all Program elements.
            var programElements = xDoc.Descendants("Programs").Elements("Program");

            // Cycle through each Program element.
            foreach (var program in programElements)
            {
                string programNameFromL5X = program.Attribute("Name")!.Value;

                // If default programName input value null, then return the XPath of the first instance of the tag if it exists.
                // OR
                // If programName input value is specified, return the XPath of the tag in that program if it exists.
                if ((programName == null) || (programName == programNameFromL5X))
                {
                    // Find all Tag elements within the current Program.
                    var programTagElements = program.Descendants("Tags").Elements("Tag");

                    // Rotate through the current Program's tag and if the tag exists, return it's XPath.
                    foreach (var tag in programTagElements)
                    {
                        if (tag.Attribute("Name")!.Value == tagName)
                        {
                            return $"Controller/Programs/Program[@Name='{programNameFromL5X}']/Tags/Tag[@Name='{tagName}']";
                        }
                    }
                }
            }

            // CONTROLLER SCOPED XPATH SEARCH
            // Find all Tag elements.
            var controllerTagElements = xDoc.Descendants("Controller").Elements("Tags").Elements("Tag");

            // Cycle through each Tag and .
            foreach (var tag in controllerTagElements)
            {
                string tagNameFromL5X = tag.Attribute("Name")!.Value;

                // Add all required parameters to the string that will populate the AOI instruction instance.
                if (tagName == tagNameFromL5X)
                {
                    return $"Controller/Tags/Tag[@Name='{tagName}']";
                }
            }

            // Print error messages to the console.
            if (programName != null)
                ConsoleMessage($"No tag found named '{tagName}' within the program '{programName}' in the file '{l5xPath}'.", "ERROR");
            else
                ConsoleMessage($"No tag found named '{tagName}' in the file '{l5xPath}'.", "ERROR");

            return null;
        }

        /// <summary>
        /// Get the Name, DataType, Usage, Required, and Visible components of each parameter from an AOI definition XML file.
        /// </summary>
        /// <param name="l5xPath">The file path to the AOI definition L5X file.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>An array of the AOIParameter structure. Each array element corresponds to a unique AOI parameter.</returns>
        private static AOIParameter[]? GetComplexParameters_FromL5X(string l5xPath, string objectName, bool printOut = false)
        {
            //// CONTROLLER SCOPED XPATH SEARCH
            //// Find all Tag elements.
            AOIParameter[]? returnDataPoints = null;
            int parameterCount;
            int paramIndex = 0;

            XDocument xDoc = XDocument.Load(l5xPath);
            var aois = xDoc.Descendants("AddOnInstructionDefinition");
            foreach (var aoi in aois)
            {
                if (aoi.Attribute("Name")!.Value == objectName)
                {
                    parameterCount = aoi.Descendants("Parameters").Elements("Parameter").Where(p => (string)p.Attribute("Usage")! != "InOut").Count();
                    returnDataPoints = new AOIParameter[parameterCount];

                    foreach (var p in aoi.Descendants("Parameter"))
                    {
                        if (p.Attribute("Usage")!.Value != "InOut")
                        {
                            returnDataPoints[paramIndex].Name = p.Attribute("Name")!.Value;
                            returnDataPoints[paramIndex].DataType = p.Attribute("DataType")!.Value;
                            returnDataPoints[paramIndex].Usage = p.Attribute("Usage")!.Value;
                            returnDataPoints[paramIndex].Required = bool.Parse(p.Attribute("Required")!.Value);
                            returnDataPoints[paramIndex].Visible = bool.Parse(p.Attribute("Visible")!.Value)!;
                            paramIndex++;
                        }
                    }

                }
            }

            var elements = xDoc.Descendants("DataTypes").Elements("DataType");
            foreach (var e in elements)
            {
                //Console.WriteLine("e.Attribute(\"Name\")!.Value: " + e.Attribute("Name")!.Value);
                if (e.Attribute("Name")!.Value == objectName)
                {
                    parameterCount = e.Descendants("Parameters").Elements("Parameter").Count();

                    returnDataPoints = new AOIParameter[parameterCount];

                    foreach (var p in e.Descendants("Parameter"))
                    {
                        if (p.Attribute("Usage")!.Value != "InOut")
                        {
                            Console.WriteLine("e.Attribute(\"Name\")!.Value: " + p.Attribute("Name")!.Value);
                            returnDataPoints[paramIndex].Name = p.Attribute("Name")!.Value;
                            returnDataPoints[paramIndex].DataType = p.Attribute("DataType")!.Value;
                            returnDataPoints[paramIndex].Usage = p.Attribute("Usage")!.Value;
                            returnDataPoints[paramIndex].Required = bool.Parse(p.Attribute("Required")!.Value);
                            returnDataPoints[paramIndex].Visible = bool.Parse(p.Attribute("Visible")!.Value)!;
                            paramIndex++;
                        }
                    }

                }
            }

            if (printOut)
                ConsoleMessage("Got complex tag sub componenet Name, Data Type, Usage, Required, & Visible values from L5X.", "STATUS");

            return returnDataPoints;
        }

        /// <summary>
        /// Convert an AOI definition L5X file to a routine L5X file with the following characteristics:<br/>
        ///  - The new routine contains 1 rung with an XIC instruction in series with a populated instance of the AOI instruction.<br/>
        ///  - The XIC instruction contains a newly created BOOL tag that is toggled depending on EnableIn logic.<br/>
        ///  - The AOI instruction has all required/visible parameters populated by a newly created AOI tag.
        /// </summary>
        /// <param name="l5xFilePath">The file path to the AOI definition L5X file.</param>
        /// <param name="routineName">The name of the new routine being created.</param>
        /// <param name="programName">The name of the program to which the routine is being added.</param>
        /// <param name="controllerName">The name of the controller to which the routine is being added.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void ConvertL5X_AOItoROUTINE(string l5xFilePath, string routineName, string programName, string controllerName, bool printOut)
        {
            string aoiName = GetAttributeValue(l5xFilePath, "AddOnInstructionDefinition", "Name", printOut)!;

            // Modify the "top" of the L5X file. This converts the definition L5X to a routine L5X.
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetName", routineName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetType", "Routine", printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "TargetSubType", "RLL", printOut);
            DeleteAttributeFromRoot(l5xFilePath, "TargetRevision", printOut);
            DeleteAttributeFromRoot(l5xFilePath, "TargetLastEdited", printOut);
            AddAttributeToComplexElement(l5xFilePath, "RSLogix5000Content", "ExportDate", DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy"), printOut);

            DeleteAttributeFromComplexElement(l5xFilePath, "AddOnInstructionDefinition", "Use", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Controller", "Name", controllerName, printOut);

            // Modify the "bottom" of the L5X file. This creates the routine, ladder logic rung, and tags as required in order to execute unit testing.
            AddElementToComplexElement(l5xFilePath, "Controller", "Tags", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tags", "Use", "Context", printOut);

            // Create the AOI tag instance using the existing XML information provided in the defintion L5X file.
            AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", "AOI_" + aoiName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", aoiName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
            string cdataforAOI_Descr = @"Automated Testing -------------------- generated AOI tag being unit tested";
            AddCDATA(l5xFilePath, "Description", cdataforAOI_Descr, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
            string cdataInfo_forData = GetAOITagCDATA_forData(l5xFilePath, aoiName, printOut)!;
            AddCDATA(l5xFilePath, "Data", cdataInfo_forData, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "Decorated", printOut);
            AddElementToComplexElement(l5xFilePath, "Data", "Structure", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Structure", "DataType", aoiName, printOut);
            List<Dictionary<string, string>> attributesList = GetDataValueMemberInfofromXML(l5xFilePath, printOut);
            AddComplexElementsWithAttributesToXml(l5xFilePath, attributesList, printOut);

            // Add all InOut tags to the Tags complex element.
            try
            {
                XDocument doc = XDocument.Load(l5xFilePath);

                // Find all "Parameter" elements.
                var parameterElements = doc.Descendants("Parameter");

                // Cycle through each AOI parameter and add it to the list if it is a required parameter.
                foreach (var param in parameterElements)
                {
                    XAttribute? nameAttribute = param.Attribute("Name");
                    XAttribute? externalAccessAttribute = param.Attribute("ExternalAccess");
                    string requiredAttributeValue = param.Attribute("Required")!.Value;

                    // Create a new tag for each string InOut parameter.
                    if ((nameAttribute != null) && (externalAccessAttribute == null) && (requiredAttributeValue == "true"))
                    {
                        AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
                        string currentInOutParamName = "AOI_" + aoiName + "_" + nameAttribute.Value;
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", currentInOutParamName, printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", "STRING", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
                        string cdataforCurrentInOutParam_Descr = $"Automated Testing -------------------- generated InOut param for AOI_{aoiName} tag";
                        AddCDATA(l5xFilePath, "Description", cdataforCurrentInOutParam_Descr, printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
                        string cdataforCurrentInOutParam_forData = (string)GetInOutParamTagCDATA_forData(l5xFilePath, nameAttribute.Value)[0];
                        AddCDATA(l5xFilePath, "Data", cdataforCurrentInOutParam_forData, printOut);
                        AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "String", printOut);
                        AddAttributeToComplexElement(l5xFilePath, "Data", "Length", "0", printOut);
                        AddCDATA(l5xFilePath, "Data", "''", printOut);
                    }
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }

            // AT_EnableIn tag
            AddElementToComplexElement(l5xFilePath, "Tags", "Tag", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Name", "AT_EnableIn", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "TagType", "Base", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "DataType", "BOOL", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Radix", "Decimal", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "Constant", "false", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Tag", "OpcUaAccess", "None", printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Description", printOut);
            string cdataforAT_EnableTag_Descr = @"Automated Testing -------------------- set the EnableIn AOI input parameter";
            AddCDATA(l5xFilePath, "Description", cdataforAT_EnableTag_Descr, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "L5K", printOut);
            string cdataforAT_EnableTag_Data = @"0";
            AddCDATA(l5xFilePath, "Data", cdataforAT_EnableTag_Data, printOut);
            AddElementToComplexElement(l5xFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Data", "Format", "Decorated", printOut);
            AddElementToComplexElement(l5xFilePath, "Data", "DataValue", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "DataType", "BOOL", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "Radix", "Decimal", printOut);
            AddAttributeToComplexElement(l5xFilePath, "DataValue", "Value", "0", printOut);

            AddElementToComplexElement(l5xFilePath, "Controller", "Programs", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Programs", "Use", "Context", printOut);

            AddElementToComplexElement(l5xFilePath, "Programs", "Program", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Program", "Use", "Context", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Program", "Name", programName, printOut);
            AddElementToComplexElement(l5xFilePath, "Program", "Routines", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routines", "Use", "Context", printOut);

            AddElementToComplexElement(l5xFilePath, "Routines", "Routine", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Use", "Target", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Name", routineName, printOut);
            AddAttributeToComplexElement(l5xFilePath, "Routine", "Type", "RLL", printOut);

            AddElementToComplexElement(l5xFilePath, "Routine", "RLLContent", printOut);
            AddElementToComplexElement(l5xFilePath, "RLLContent", "Rung", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Rung", "Number", "0", printOut);
            AddAttributeToComplexElement(l5xFilePath, "Rung", "Type", "N", printOut);

            AddElementToComplexElement(l5xFilePath, "Rung", "Comment", printOut);
            string cdataInfoforRung0Comment = @"AUTOMATED TESTING | " + aoiName + @" AOI UNIT TEST
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
This a programmatically created rung with a populated instance of the AOI instruction added using the Logix Designer SDK.";
            AddCDATA(l5xFilePath, "Comment", cdataInfoforRung0Comment, printOut);

            AddElementToComplexElement(l5xFilePath, "Rung", "Text", printOut);
            string cdataInfo_forText = GetCDATAfromXML_forText(l5xFilePath, printOut);
            AddCDATA(l5xFilePath, "Text", cdataInfo_forText, printOut);
        }

        /// <summary>
        /// Get the CDATA contents needed to create a Studio 5000 Logix Designer Tag for an AOI InOut parameter.
        /// </summary>
        /// <param name="l5xFilePath">The AOI definition L5X file path.</param>
        /// <param name="paramName">The name of the target parameter.</param>
        /// <returns>A string of CDATA contents needed to create a Studio 5000 Logix Designer tag.</returns>
        /// <exception cref="ArgumentException">Thrown if integers for tag dimension size cannot be parsed properly from L5X file.</exception>
        public static object[] GetInOutParamTagCDATA_forData(string l5xFilePath, string paramName, int boolCount = 0, bool printOut = false)
        {
            // The variable returnCDATA contains the final result to be returned.
            StringBuilder returnCDATA = new();
            object[] returnObjectArray = new object[2];

            // Variables required to account for tag nesting.
            StringBuilder nested0Contents = new();
            StringBuilder nested1Contents = new();
            StringBuilder nested2Contents = new();
            StringBuilder nested3Contents = new();
            StringBuilder nested4Contents = new();
            StringBuilder nested5Contents = new();
            StringBuilder nested6Contents = new();
            StringBuilder nested7Contents = new();
            StringBuilder nested8Contents = new();
            string nested0DataType = "";
            string nested1DataType = "";
            string nested2DataType = "";
            string nested3DataType = "";
            string nested4DataType = "";
            string nested5DataType = "";
            string nested6DataType = "";
            string nested7DataType = "";
            string nested8DataType = "";
            int nested0Dimensions = 0;
            int nested1Dimensions = 0;
            int nested2Dimensions = 0;
            int nested3Dimensions = 0;
            int nested4Dimensions = 0;
            int nested5Dimensions = 0;
            int nested6Dimensions = 0;
            int nested7Dimensions = 0;
            int nested8Dimensions = 0;

            XDocument xdoc = XDocument.Load(l5xFilePath);

            // Find all "Parameter" elements from the AOI definition L5X file.
            var nested0ParameterElements = xdoc.Descendants("Parameter");
            var nested0LocalTagElements = xdoc.Descendants("LocalTag");

            // Cycle through each AOI parameter and get the nested level 0 target datatype and dimensions.
            foreach (var n0pe in nested0ParameterElements)
            {
                if (paramName == n0pe.Attribute("Name")!.Value)
                {
                    nested0DataType = n0pe.Attribute("DataType")!.Value;
                    if (n0pe.Attribute("Dimensions") != null)
                    {
                        if (!int.TryParse(n0pe.Attribute("Dimensions")!.Value, out nested0Dimensions))
                            throw new ArgumentException("Parameter dimensions must be a valid integer.");
                    }
                    break;
                }
                else
                {
                    foreach (var n0lte in nested0LocalTagElements)
                    {
                        if (paramName == n0lte.Attribute("Name")!.Value)
                        {
                            nested0DataType = n0lte.Attribute("DataType")!.Value;
                            if (n0lte.Attribute("Dimensions") != null)
                            {
                                if (!int.TryParse(n0lte.Attribute("Dimensions")!.Value, out nested0Dimensions))
                                    throw new ArgumentException("Parameter dimensions must be a valid integer.");
                            }
                            break;
                        }
                    }
                }
            }

            string? nested0CDATA = GetS5kAtomicTagCDATA_forData(nested0DataType, nested0Dimensions > 0);

            // If atomic data not in an array.
            if ((nested0CDATA != null) && (nested0Dimensions == 0))
            {
                if ((nested0DataType == "BOOL") || (nested0DataType == "BIT"))
                {
                    boolCount++;
                    if ((boolCount & 31) == 1)
                        returnCDATA.Append(nested0CDATA);
                }
                else
                    returnCDATA.Append(nested0CDATA);
            }
            // If atomic data in an array.
            else if ((nested0CDATA != null) && (nested0Dimensions > 0))
            {
                nested0Contents.Append(CreateArrayCDATA(nested0CDATA, nested0Dimensions));
                returnCDATA.Append(nested0Contents.ToString());
            }
            // If data is complex.
            else if (nested0CDATA == null)
            {
                #region NESTED LEVELS: 1 to 8
                // Get the non-hidden members of the target DataType complex element.
                var nested1DataTypeMemberElements = xdoc
                    .Descendants("DataTypes")
                    .Elements("DataType")
                    .FirstOrDefault(dt => dt.Attribute("Name")!.Value == nested0DataType)!
                    .Descendants("Members")
                    .Elements("Member")
                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                // Rotate through the level 1 nested member elements & add them to the CDATA stringbuilder.
                nested1Contents.Append('[');
                foreach (var n1dtme in nested1DataTypeMemberElements)
                {
                    // Get the data type of the current nested level and the dimension of the current member.
                    nested1DataType = n1dtme.Attribute("DataType")!.Value;
                    if (n1dtme.Attribute("Dimension") != null)
                    {
                        if (!int.TryParse(n1dtme.Attribute("Dimension")!.Value, out nested1Dimensions))
                            throw new ArgumentException("Tag dimensions (at 1 nested level) must be a valid integer.");
                    }

                    string? nested1CDATA = GetS5kAtomicTagCDATA_forData(nested1DataType, nested1Dimensions > 0);

                    if ((nested1CDATA != null) && (nested1Dimensions == 0))
                    {
                        if ((nested1DataType == "BOOL") || (nested1DataType == "BIT"))
                        {
                            boolCount++;
                            if ((boolCount & 31) == 1)
                                nested1Contents.Append(nested1CDATA);
                        }
                        else
                            nested1Contents.Append(nested1CDATA);
                    }
                    else if ((nested1CDATA != null) && (nested1Dimensions > 0))
                    {
                        nested1Contents.Append(CreateArrayCDATA(nested1CDATA, nested1Dimensions));
                    }
                    else if (nested1CDATA == null)
                    {
                        #region NESTED LEVELS: 2 to 8
                        var nested2DataTypeMemberElements = xdoc
                            .Descendants("DataTypes")
                            .Elements("DataType")
                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested1DataType)!
                            .Descendants("Members")
                            .Elements("Member")
                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                        // Rotate through the level 2 nested member elements & add them to the CDATA stringbuilder.
                        nested2Contents.Append('[');
                        foreach (var n2dtme in nested2DataTypeMemberElements)
                        {
                            nested2DataType = n2dtme.Attribute("DataType")!.Value;
                            if (n2dtme.Attribute("Dimension") != null)
                            {
                                if (!int.TryParse(n2dtme.Attribute("Dimension")!.Value, out nested2Dimensions))
                                    throw new ArgumentException("Tag dimensions (at 2 nested level) must be a valid integer.");
                            }

                            string? nested2CDATA = GetS5kAtomicTagCDATA_forData(nested2DataType, nested2Dimensions > 0);

                            if ((nested2CDATA != null) && (nested2Dimensions == 0))
                            {
                                if ((nested2DataType == "BOOL") || (nested2DataType == "BIT"))
                                {
                                    boolCount++;
                                    if ((boolCount & 31) == 1)
                                        nested2Contents.Append(nested2CDATA);
                                }
                                else
                                    nested2Contents.Append(nested2CDATA);
                            }
                            else if ((nested2CDATA != null) && (nested2Dimensions > 0))
                            {
                                nested2Contents.Append(CreateArrayCDATA(nested2CDATA, nested2Dimensions));
                            }
                            else if (nested2CDATA == null)
                            {
                                #region NESTED LEVELS: 3 to 8
                                var nested3DataTypeMemberElements = xdoc
                                    .Descendants("DataTypes")
                                    .Elements("DataType")
                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested2DataType)!
                                    .Descendants("Members")
                                    .Elements("Member")
                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                // Rotate through the level 3 nested member elements & add them to the CDATA stringbuilder.
                                nested3Contents.Append('[');
                                foreach (var n3dtme in nested3DataTypeMemberElements)
                                {
                                    nested3DataType = n3dtme.Attribute("DataType")!.Value;
                                    if (n3dtme.Attribute("Dimension") != null)
                                    {
                                        if (!int.TryParse(n3dtme.Attribute("Dimension")!.Value, out nested3Dimensions))
                                            throw new ArgumentException("Tag dimensions (at 3 nested level) must be a valid integer.");
                                    }

                                    string? nested3CDATA = GetS5kAtomicTagCDATA_forData(nested3DataType, nested3Dimensions > 0);

                                    if ((nested3CDATA != null) && (nested3Dimensions == 0))
                                    {
                                        if ((nested3DataType == "BOOL") || (nested3DataType == "BIT"))
                                        {
                                            boolCount++;
                                            if ((boolCount & 31) == 1)
                                                nested3Contents.Append(nested3CDATA);
                                        }
                                        else
                                            nested3Contents.Append(nested3CDATA);
                                    }
                                    else if ((nested3CDATA != null) && (nested3Dimensions > 0))
                                    {
                                        nested3Contents.Append(CreateArrayCDATA(nested3CDATA, nested3Dimensions));
                                    }
                                    else if (nested3CDATA == null)
                                    {
                                        #region NESTED LEVELS: 4 to 8
                                        var nested4DataTypeMemberElements = xdoc
                                            .Descendants("DataTypes")
                                            .Elements("DataType")
                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested3DataType)!
                                            .Descendants("Members")
                                            .Elements("Member")
                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                        // Rotate through the level 4 nested member elements & add them to the CDATA stringbuilder.
                                        nested4Contents.Append('[');
                                        foreach (var n4dtme in nested4DataTypeMemberElements)
                                        {
                                            nested4DataType = n4dtme.Attribute("DataType")!.Value;
                                            if (n4dtme.Attribute("Dimension") != null)
                                            {
                                                if (!int.TryParse(n4dtme.Attribute("Dimension")!.Value, out nested4Dimensions))
                                                    throw new ArgumentException("Tag dimensions (at 4 nested level) must be a valid integer.");
                                            }

                                            string? nested4CDATA = GetS5kAtomicTagCDATA_forData(nested4DataType, nested4Dimensions > 0);

                                            if ((nested4CDATA != null) && (nested4Dimensions == 0))
                                            {
                                                if ((nested4DataType == "BOOL") || (nested4DataType == "BIT"))
                                                {
                                                    boolCount++;
                                                    if ((boolCount & 31) == 1)
                                                        nested4Contents.Append(nested4CDATA);
                                                }
                                                else
                                                    nested4Contents.Append(nested4CDATA);
                                            }
                                            else if ((nested4CDATA != null) && (nested4Dimensions > 0))
                                            {
                                                nested4Contents.Append(CreateArrayCDATA(nested4CDATA, nested4Dimensions));
                                            }
                                            else if (nested4CDATA == null)
                                            {
                                                #region NESTED LEVELS: 5 to 8
                                                var nested5DataTypeMemberElements = xdoc
                                                    .Descendants("DataTypes")
                                                    .Elements("DataType")
                                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested4DataType)!
                                                    .Descendants("Members")
                                                    .Elements("Member")
                                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                // Rotate through the level 5 nested member elements & add them to the CDATA stringbuilder.
                                                nested5Contents.Append('[');
                                                foreach (var n5dtme in nested5DataTypeMemberElements)
                                                {
                                                    nested5DataType = n5dtme.Attribute("DataType")!.Value;
                                                    if (n5dtme.Attribute("Dimension") != null)
                                                    {
                                                        if (!int.TryParse(n5dtme.Attribute("Dimension")!.Value, out nested5Dimensions))
                                                            throw new ArgumentException("Tag dimensions (at 5 nested level) must be a valid integer.");
                                                    }

                                                    string? nested5CDATA = GetS5kAtomicTagCDATA_forData(nested5DataType, nested5Dimensions > 0);

                                                    if ((nested5CDATA != null) && (nested5Dimensions == 0))
                                                    {
                                                        if ((nested5DataType == "BOOL") || (nested5DataType == "BIT"))
                                                        {
                                                            boolCount++;
                                                            if ((boolCount & 31) == 1)
                                                                nested5Contents.Append(nested5CDATA);
                                                        }
                                                        else
                                                            nested5Contents.Append(nested5CDATA);
                                                    }
                                                    else if ((nested5CDATA != null) && (nested5Dimensions > 0))
                                                    {
                                                        nested5Contents.Append(CreateArrayCDATA(nested5CDATA, nested5Dimensions));
                                                    }
                                                    else if (nested5CDATA == null)
                                                    {
                                                        #region NESTED LEVELS: 6 to 8
                                                        var nested6DataTypeMemberElements = xdoc
                                                            .Descendants("DataTypes")
                                                            .Elements("DataType")
                                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested5DataType)!
                                                            .Descendants("Members")
                                                            .Elements("Member")
                                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                        // Rotate through the level 6 nested member elements & add them to the CDATA stringbuilder.
                                                        nested6Contents.Append('[');
                                                        foreach (var n6dtme in nested6DataTypeMemberElements)
                                                        {
                                                            nested6DataType = n6dtme.Attribute("DataType")!.Value;
                                                            if (n6dtme.Attribute("Dimension") != null)
                                                            {
                                                                if (!int.TryParse(n6dtme.Attribute("Dimension")!.Value, out nested6Dimensions))
                                                                    throw new ArgumentException("Tag dimensions (at 6 nested level) must be a valid integer.");
                                                            }

                                                            string? nested6CDATA = GetS5kAtomicTagCDATA_forData(nested6DataType, nested6Dimensions > 0);

                                                            if ((nested6CDATA != null) && (nested6Dimensions == 0))
                                                            {
                                                                if ((nested6DataType == "BOOL") || (nested6DataType == "BIT"))
                                                                {
                                                                    boolCount++;
                                                                    if ((boolCount & 31) == 1)
                                                                        nested6Contents.Append(nested6CDATA);
                                                                }
                                                                else
                                                                    nested6Contents.Append(nested6CDATA);
                                                            }
                                                            else if ((nested6CDATA != null) && (nested6Dimensions > 0))
                                                            {
                                                                nested6Contents.Append(CreateArrayCDATA(nested6CDATA, nested6Dimensions));
                                                            }
                                                            else if (nested6CDATA == null)
                                                            {
                                                                #region NESTED LEVELS: 7 to 8
                                                                var nested7DataTypeMemberElements = xdoc
                                                                    .Descendants("DataTypes")
                                                                    .Elements("DataType")
                                                                    .FirstOrDefault(e => (string)e.Attribute("Name")! == nested6DataType)!
                                                                    .Descendants("Members")
                                                                    .Elements("Member")
                                                                    .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                                // Rotate through the level 7 nested member elements & add them to the CDATA stringbuilder.
                                                                nested7Contents.Append('[');
                                                                foreach (var n7dtme in nested7DataTypeMemberElements)
                                                                {
                                                                    nested7DataType = n7dtme.Attribute("DataType")!.Value;
                                                                    if (n7dtme.Attribute("Dimension") != null)
                                                                    {
                                                                        if (!int.TryParse(n7dtme.Attribute("Dimension")!.Value, out nested7Dimensions))
                                                                            throw new ArgumentException("Tag dimensions (at 7 nested level) must be a valid integer.");
                                                                    }

                                                                    string? nested7CDATA = GetS5kAtomicTagCDATA_forData(nested7DataType, nested7Dimensions > 0);

                                                                    if ((nested7CDATA != null) && (nested7Dimensions == 0))
                                                                    {
                                                                        if ((nested7DataType == "BOOL") || (nested7DataType == "BIT"))
                                                                        {
                                                                            boolCount++;
                                                                            if ((boolCount & 31) == 1)
                                                                                nested7Contents.Append(nested7CDATA);
                                                                        }
                                                                        else
                                                                            nested7Contents.Append(nested7CDATA);
                                                                    }
                                                                    else if ((nested7CDATA != null) && (nested7Dimensions > 0))
                                                                    {
                                                                        nested7Contents.Append(CreateArrayCDATA(nested7CDATA, nested7Dimensions));
                                                                    }
                                                                    else if (nested7CDATA == null)
                                                                    {
                                                                        #region NESTED LEVELS: 8 to 8
                                                                        var nested8DataTypeMemberElements = xdoc
                                                                            .Descendants("DataTypes")
                                                                            .Elements("DataType")
                                                                            .FirstOrDefault(e => (string)e.Attribute("Name")! == nested7DataType)!
                                                                            .Descendants("Members")
                                                                            .Elements("Member")
                                                                            .Where(m => m.Attribute("Hidden")!.Value == "false");

                                                                        // Rotate through the level 8 nested member elements & add them to the CDATA stringbuilder.
                                                                        nested8Contents.Append('[');
                                                                        foreach (var n8dtme in nested8DataTypeMemberElements)
                                                                        {
                                                                            nested8DataType = n8dtme.Attribute("DataType")!.Value;
                                                                            if (n8dtme.Attribute("Dimension") != null)
                                                                            {
                                                                                if (!int.TryParse(n8dtme.Attribute("Dimension")!.Value, out nested8Dimensions))
                                                                                    throw new ArgumentException("Tag dimensions (at 8 nested level) must be a valid integer.");
                                                                            }

                                                                            string? nested8CDATA = GetS5kAtomicTagCDATA_forData(nested8DataType, nested8Dimensions > 0);

                                                                            if ((nested8CDATA != null) && (nested8Dimensions == 0))
                                                                            {
                                                                                if ((nested8DataType == "BOOL") || (nested8DataType == "BIT"))
                                                                                {
                                                                                    boolCount++;
                                                                                    if ((boolCount & 31) == 1)
                                                                                        nested8Contents.Append(nested8CDATA);
                                                                                }
                                                                                else
                                                                                    nested8Contents.Append(nested8CDATA);
                                                                            }
                                                                            else if ((nested8CDATA != null) && (nested8Dimensions > 0))
                                                                            {
                                                                                nested8Contents.Append(CreateArrayCDATA(nested8CDATA, nested8Dimensions));
                                                                            }
                                                                            else if (nested8CDATA == null)
                                                                            {
                                                                                ConsoleMessage("Nesting tags beyond 8 levels not supported.", "ERROR");
                                                                            }

                                                                            // Add a comma for all situations other than boolean "bit packing".
                                                                            if ((boolCount & 31) == 1)
                                                                                nested8Contents.Append(',');
                                                                            else if (nested8DataType != "BIT")
                                                                                nested8Contents.Append(',');
                                                                        }
                                                                        nested8Contents.Length--;    // Remove the last apostrophe character ','
                                                                        nested8Contents.Append(']'); // Finish the nested component syntax.

                                                                        if (nested7Dimensions == 0)
                                                                        {
                                                                            nested7Contents.Append(nested8Contents.ToString());
                                                                        }
                                                                        else if (nested7Dimensions > 0)
                                                                        {
                                                                            nested7Contents.Append(CreateArrayCDATA(nested8Contents.ToString(), nested7Dimensions));
                                                                        }

                                                                        nested8Contents.Clear();
                                                                        #endregion
                                                                    }

                                                                    // Add a comma for all situations other than boolean "bit packing".
                                                                    if ((boolCount & 31) == 1)
                                                                        nested7Contents.Append(',');
                                                                    else if (nested7DataType != "BIT")
                                                                        nested7Contents.Append(',');
                                                                }
                                                                nested7Contents.Length--;    // Remove the last apostrophe character ','
                                                                nested7Contents.Append(']'); // Finish the nested component syntax.

                                                                if (nested6Dimensions == 0)
                                                                {
                                                                    nested6Contents.Append(nested7Contents.ToString());
                                                                }
                                                                else if (nested6Dimensions > 0)
                                                                {
                                                                    nested6Contents.Append(CreateArrayCDATA(nested7Contents.ToString(), nested6Dimensions));
                                                                }

                                                                nested7Contents.Clear();
                                                                #endregion
                                                            }

                                                            // Add a comma for all situations other than boolean "bit packing".
                                                            if ((boolCount & 31) == 1)
                                                                nested6Contents.Append(',');
                                                            else if (nested6DataType != "BIT")
                                                                nested6Contents.Append(',');
                                                        }
                                                        nested6Contents.Length--;    // Remove the last apostrophe character ','
                                                        nested6Contents.Append(']'); // Finish the nested component syntax.

                                                        if (nested5Dimensions == 0)
                                                        {
                                                            nested5Contents.Append(nested6Contents.ToString());
                                                        }
                                                        else if (nested5Dimensions > 0)
                                                        {
                                                            nested5Contents.Append(CreateArrayCDATA(nested6Contents.ToString(), nested5Dimensions));
                                                        }

                                                        nested6Contents.Clear();
                                                        #endregion
                                                    }

                                                    // Add a comma for all situations other than boolean "bit packing".
                                                    if ((boolCount & 31) == 1)
                                                        nested5Contents.Append(',');
                                                    else if (nested5DataType != "BIT")
                                                        nested5Contents.Append(',');
                                                }
                                                nested5Contents.Length--;    // Remove the last apostrophe character ','
                                                nested5Contents.Append(']'); // Finish the nested component syntax.

                                                if (nested4Dimensions == 0)
                                                {
                                                    nested4Contents.Append(nested5Contents.ToString());
                                                }
                                                else if (nested4Dimensions > 0)
                                                {
                                                    nested4Contents.Append(CreateArrayCDATA(nested5Contents.ToString(), nested4Dimensions));
                                                }

                                                nested5Contents.Clear();
                                                #endregion
                                            }

                                            // Add a comma for all situations other than boolean "bit packing".
                                            if ((boolCount & 31) == 1)
                                                nested4Contents.Append(',');
                                            else if (nested4DataType != "BIT")
                                                nested4Contents.Append(',');
                                        }
                                        nested4Contents.Length--;    // Remove the last apostrophe character ','
                                        nested4Contents.Append(']'); // Finish the nested component syntax.

                                        if (nested3Dimensions == 0)
                                        {
                                            nested3Contents.Append(nested4Contents.ToString());
                                        }
                                        else if (nested3Dimensions > 0)
                                        {
                                            nested3Contents.Append(CreateArrayCDATA(nested4Contents.ToString(), nested3Dimensions));
                                        }

                                        nested4Contents.Clear();
                                        #endregion

                                    }

                                    // Add a comma for all situations other than boolean "bit packing".
                                    if ((boolCount & 31) == 1)
                                        nested3Contents.Append(',');
                                    else if (nested3DataType != "BIT")
                                        nested3Contents.Append(',');
                                }
                                nested3Contents.Length--;    // Remove the last apostrophe character ','
                                nested3Contents.Append(']'); // Finish the nested component syntax.

                                if (nested2Dimensions == 0)
                                {
                                    nested2Contents.Append(nested3Contents.ToString());
                                }
                                else if (nested2Dimensions > 0)
                                {
                                    nested2Contents.Append(CreateArrayCDATA(nested3Contents.ToString(), nested2Dimensions));
                                }

                                nested3Contents.Clear();
                                #endregion
                            }

                            // Add a comma for all situations other than boolean "bit packing".
                            if ((boolCount & 31) == 1)
                                nested2Contents.Append(',');
                            else if (nested2DataType != "BIT")
                                nested2Contents.Append(',');
                        }
                        nested2Contents.Length--;    // Remove the last apostrophe character ','
                        nested2Contents.Append(']'); // Finish the nested component syntax.

                        if (nested1Dimensions == 0)
                        {
                            nested1Contents.Append(nested2Contents.ToString());
                        }
                        else if (nested1Dimensions > 0)
                        {
                            nested1Contents.Append(CreateArrayCDATA(nested2Contents.ToString(), nested1Dimensions));
                        }

                        nested2Contents.Clear();
                        #endregion
                    }

                    // Add a comma for all situations other than boolean "bit packing".
                    if ((boolCount & 31) == 1)
                        nested1Contents.Append(',');
                    else if (nested1DataType != "BIT")
                        nested1Contents.Append(',');
                }
                nested1Contents.Length--;    // Remove the last apostrophe character ','
                nested1Contents.Append(']'); // Finish the nested component syntax.

                if (nested0Dimensions == 0)
                {
                    returnCDATA.Append(nested1Contents.ToString());
                }
                else if (nested0Dimensions > 0)
                {
                    returnCDATA.Append(CreateArrayCDATA(nested1Contents.ToString(), nested0Dimensions));
                }

                nested1Contents.Clear();
                #endregion
            }

            if (printOut)
                ConsoleMessage($"CDATA contents: {returnCDATA.ToString()}", "STATUS");

            returnObjectArray[0] = returnCDATA.ToString();
            returnObjectArray[1] = boolCount;
            return returnObjectArray;
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the 'Data' complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string? GetAOITagCDATA_forData(string xmlFilePath, string aoiName, bool printOut)
        {
            int boolCount = 0;
            List<string> paramElemStringList = new List<string>();
            List<string> localTagElemStringList = new List<string>();

            //try
            //{
            XDocument doc = XDocument.Load(xmlFilePath);

            // Get a list filtered to contain only CDATA information from nonboolean "Parameter" elements.
            var parameterElements = doc
                .Descendants("AddOnInstructionDefinition")
                .Where(param => param.Attribute("Name")?.Value == aoiName)
                .Descendants("Parameters")
                .Elements("Parameter")
                .Where(param => param.Attribute("Usage")?.Value != "InOut");


            foreach (var pe in parameterElements)
            {
                object[] currentParamCDATA = GetInOutParamTagCDATA_forData(xmlFilePath, pe.Attribute("Name")!.Value, boolCount);
                string currentCDATA = (string)currentParamCDATA[0];
                boolCount = 0;
                boolCount += (int)currentParamCDATA[1];

                if (currentCDATA != "")
                    paramElemStringList.Add(currentCDATA);
            }

            // Join all parameterElements list elements into a single string, with each element separated by a comma without spaces.
            string joined_pCDATA = string.Join(",", paramElemStringList);

            // Get a list filtered to contain only CDATA information from nonboolean "LocalTag" elements.
            var localtagElements = doc
                .Descendants("AddOnInstructionDefinition")
                .Where(param => param.Attribute("Name")?.Value == aoiName)
                .Descendants("LocalTags")
                .Elements("LocalTag")
                .Where(param => param.Attribute("Usage")?.Value != "InOut");

            foreach (var lte in localtagElements)
            {
                object[] currentLocalTagCDATA = GetInOutParamTagCDATA_forData(xmlFilePath, lte.Attribute("Name")!.Value, boolCount);
                string currentCDATA = (string)currentLocalTagCDATA[0];
                boolCount = 0;
                boolCount += (int)currentLocalTagCDATA[1];

                if (currentCDATA != "")
                    localTagElemStringList.Add(currentCDATA);
            }

            // Join all localtagElements list elements into a single string, with each element separated by a comma without spaces.
            string joined_ltCDATA = string.Join(",", localTagElemStringList);

            // If no local tags, don't include the joined local tag CDATA information.
            string returnString;
            if (joined_ltCDATA == "")
            {
                returnString = "[" + joined_pCDATA + "]";
            }
            else
            {
                returnString = "[" + joined_pCDATA + "," + joined_ltCDATA + "]";
            }

            // Create the final formatted string to be used as CDATA content information (in the Data complex element of L5X).

            if (printOut)
                ConsoleMessage($"CDATA contents: {returnString}", "STATUS");

            return returnString;
            //}
            //catch (Exception e)
            //{
            //    ConsoleMessage("GetAOITagCDATA_forData Method: " + e.Message, "ERROR");
            //    return null;
            //}
        }

        /// <summary>
        /// Get the value of an attribute for a specific complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element containing the attribute that will have its value returned.</param>
        /// <param name="attributeName">The name of the attribute that will have its value returned.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>The value of an attribute for a specific complex element, or null if the attribute is not found.</returns>
        public static string? GetAttributeValue(string xmlFilePath, string complexElementName, string attributeName, bool printOut)
        {
            XDocument xdoc = XDocument.Load(xmlFilePath);
            XElement? complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

            if (complexElement != null)
            {
                XAttribute? attribute = complexElement.Attribute(attributeName);
                if (attribute != null)
                {
                    return attribute.Value;
                }
                else if (printOut)
                {
                    ConsoleMessage($"Attribute '{attributeName}' not found in element '{complexElementName}'.", "ERROR");
                }
            }
            else if (printOut)
            {
                ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
            }

            return null;
        }

        /// <summary>
        /// Delete an attribute (name & value) from a complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element containing the attribute to be deleted.</param>
        /// <param name="attributeToDelete">The name of the attribute to be deleted.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void DeleteAttributeFromComplexElement(string xmlFilePath, string complexElementName, string attributeToDelete, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).FirstOrDefault()!;

                if (complexElement != null)
                {
                    XAttribute attribute = complexElement.Attribute(attributeToDelete)!;

                    if (attribute != null)
                    {
                        attribute.Remove();

                        if (printOut)
                        {
                            ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the element '{complexElementName}'.", "STATUS");
                        }

                        xdoc.Save(xmlFilePath);
                    }
                    else if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeToDelete}' not found in element '{complexElementName}'.", "ERROR");
                    }
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Delete an attribute from the root complex element of a 
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="attributeToDelete">The name of the attribute to be deleted.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void DeleteAttributeFromRoot(string xmlFilePath, string attributeToDelete, bool printOut)
        {
            XDocument xdoc = XDocument.Load(xmlFilePath);
            XElement root = xdoc.Root!;
            XAttribute attribute = root.Attribute(attributeToDelete)!;

            if (attribute != null)
            {
                attribute.Remove();

                if (printOut)
                {
                    string complexElementName = "RSLogix5000Content";
                    ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the root complex element '{complexElementName}'.", "STATUS");
                }

                xdoc.Save(xmlFilePath);
            }
            else if (printOut)
            {
                ConsoleMessage($"Attribute '{attributeToDelete}' not found in the root complex element.", "ERROR");
            }
        }

        /// <summary>
        /// Add (or overwrite) an attribute name and value to a complex element in XML.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the attribute will be added.</param>
        /// <param name="attributeName">The name of the attribute to be added.</param>
        /// <param name="attributeValue">The value of the attirbute to be added.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void AddAttributeToComplexElement(string xmlFilePath, string complexElementName, string attributeName, string attributeValue,
            bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    complexElement.SetAttributeValue(attributeName, attributeValue);

                    if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element " +
                            $"'{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Add an element to a complex element in XML.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the element will be added.</param>
        /// <param name="newElementName">The name of the new element.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        public static void AddElementToComplexElement(string xmlFilePath, string complexElementName, string newElementName, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    XElement newElement = new XElement(newElementName);
                    complexElement.Add(newElement);

                    if (printOut)
                    {
                        ConsoleMessage($"Element '{newElementName}' has been added to the complex element '{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Create a new CDATA element to the last or default instance of a specified complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the CDATA element will be added.</param>
        /// <param name="cdataContent">The contents of the CDATA element.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void AddCDATA(string xmlFilePath, string complexElementName, string cdataContent, bool printOut)
        {
            try
            {
                XDocument xdoc = XDocument.Load(xmlFilePath);
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault()!;

                if (complexElement != null)
                {
                    XCData cdataSection = new XCData(cdataContent);
                    complexElement.Add(cdataSection);

                    if (printOut)
                    {
                        ConsoleMessage($"A new CDATA section has been created and added to the element '{complexElementName}'.", "STATUS");
                    }

                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the Text complex element.<br/>
        /// This method programmatically gathers and formats where the information needed for a new AOI tag.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string GetCDATAfromXML_forText(string xmlFilePath, bool printOut)
        {
            // The name of the AOI being tested.
            string? aoiName = GetAttributeValue(xmlFilePath, "AddOnInstructionDefinition", "Name", printOut);

            // Initialize the StringBuilder that will contain the AOI parameter tag names.
            StringBuilder aoiTagParameterNames = new();

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);

                // Find all "Parameter" elements.
                var parameterElements = doc.Descendants("Parameter");

                // Cycle through each AOI parameter and add it to the list if it is a required parameter.
                foreach (var param in parameterElements)
                {
                    XAttribute? nameAttribute = param.Attribute("Name");
                    XAttribute? externalAccessAttribute = param.Attribute("ExternalAccess");
                    string requiredAttributeValue = param.Attribute("Required")!.Value;

                    // Add all required parameters to the string that will populate the AOI instruction instance.
                    if ((nameAttribute != null) && (requiredAttributeValue == "true"))
                    {
                        if (externalAccessAttribute != null) // In/Out Params: no InOut because they do not have the ExternalAccess attribute.
                        {
                            aoiTagParameterNames.Append($",AOI_{aoiName}.{nameAttribute.Value}");
                        }
                        else // InOut parameters formatted such that they're not part of the AOI tag.
                        {
                            aoiTagParameterNames.Append($",AOI_{aoiName}_{nameAttribute.Value}");
                        }
                    }
                }

                string returnString = $"XIC(AT_EnableIn){aoiName}(AOI_{aoiName}{aoiTagParameterNames});";

                if (printOut)
                {
                    ConsoleMessage($"CDATA contents: {returnString}", "STATUS");
                }

                return returnString;
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
                return e.Message;
            }
        }

        /// <summary>
        /// Get all the attribute names and values for each parameter in an AOI L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A list of dictionaries for each AOI parameter's attributes.</returns>
        public static List<Dictionary<string, string>> GetDataValueMemberInfofromXML(string xmlFilePath, bool printOut)
        {
            List<Dictionary<string, string>> return_attributeList = new List<Dictionary<string, string>>();

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);

                // Cycle through each "Parameter" element in the L5X file.
                foreach (var parameterElement in doc.Descendants("Parameter"))
                {
                    if (parameterElement.Attribute("Radix") != null)
                    {
                        Dictionary<string, string> attributes = new Dictionary<string, string>
                        {
                            { "Name", parameterElement.Attribute("Name")!.Value },
                            { "DataType", parameterElement.Attribute("DataType")!.Value },
                            { "Radix", parameterElement.Attribute("Radix")!.Value }
                        };

                        // Store the new dictionary containing attributes for a single AOI parameter.
                        return_attributeList.Add(attributes);
                    }
                }

                if (printOut)
                    ConsoleMessage($"Got element information from L5X (Name, Data Type, & Radix) per AOI parameter.", "STATUS");
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }

            return return_attributeList;
        }

        /// <summary>
        /// For each AOI parameter, add the element "DataValueMember" with its attributes to the L5X complex element "Structure".<br/>
        /// This method creates XML children needed to create an AOI tag in the L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="attributesList">A list of dictionaries for each AOI parameter's attributes.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void AddComplexElementsWithAttributesToXml(string xmlFilePath, List<Dictionary<string, string>> attributesList, bool printOut)
        {
            try
            {
                foreach (var attributes in attributesList)
                {
                    // Add a new element "DataValueMember" to complex element "Structure" for each AOI parameter.
                    AddElementToComplexElement(xmlFilePath, "Structure", "DataValueMember", printOut);

                    // Add the "Name" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Name", attributes["Name"], printOut);

                    // Add the "DataType" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "DataType", attributes["DataType"], printOut);

                    // Add the "Radix" attribute and its value for the current AOI parameter.
                    // Note: BOOL datatype parameters don't have a "Radix" attribute and are therefore skipped.
                    if (attributes["DataType"] != "BOOL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Radix", attributes["Radix"], printOut);
                    }

                    // Add the "Value" attribute and its value for the current AOI parameter.
                    // Note: For AOIs, the only BOOL parameter with a value of 1 is "EnableIn".
                    // Note: For REAL datatype parameters, their intial zero value has the notation "0.0". All else is "0".
                    if (attributes["Name"] == "EnableIn")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "1", printOut);
                    }
                    else if (attributes["DataType"] == "REAL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0.0", printOut);
                    }
                    else
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0", printOut);
                    }
                }
                if (printOut)
                {
                    ConsoleMessage("Complex elements added.", "STATUS");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }
        #endregion

        #region METHODS: get excel file information
        /// <summary>
        /// In the first worksheet of an Excel workbook, get the number of populated cells in the specified row after a specified column.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="rowNumber">The row in which the populated cell count is derived.</param>
        /// <param name="columnNumber">The column number after which the method begins counting populated cells.</param>
        /// <returns>The number of populated cells in the specified row.</returns>
        private static int GetNumberOfTestCases(string excelFilePath, int rowNumber, int columnNumber)
        {
            int returnCellCount = 0;
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int maxColumnNum = worksheet.Dimension.End.Column;

                for (int col = columnNumber; col <= maxColumnNum; col++)
                {
                    var cellValue = worksheet.Cells[rowNumber, col].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        returnCellCount++;
                }
            }
            return returnCellCount;
        }

        /// <summary>
        /// In the first worksheet of an Excel workbook, get the number of populated cells in the specified column.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="columnNumber">The column in which the populated cell count is derived.</param>
        /// <returns>The number of populated cells in the specified column.</returns>
        private static int GetPopulatedCellsInColumnCount(string excelFilePath, int columnNumber)
        {
            int returnCellCount = 0;
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int maxRowNum = worksheet.Dimension.End.Row;

                for (int row = 1; row <= maxRowNum; row++)
                {
                    var cellValue = worksheet.Cells[row, columnNumber].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        returnCellCount++;
                }
            }
            return returnCellCount;
        }

        /// <summary>
        /// Collect the values of each AOI parameter to be used during a test case from a specifically formatted excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path of the excel workbook containing the test case information.</param>
        /// <param name="columnNumber">The column number of a test case in the excel file.</param>
        /// <returns>A dictionary where the Key is an AOI parameter name and the Value is an AOI parameter value.</returns>
        public static Dictionary<string, string> GetExcelTestCaseValues(string excelFilePath, int columnNumber)
        {
            Dictionary<string, string> returnDictionary = [];

            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int numberOfParameters = GetPopulatedCellsInColumnCount(excelFilePath, 2) - 6;

                for (int rowNumber = 19; rowNumber < (numberOfParameters + 19); rowNumber++)
                {
                    returnDictionary[worksheet.Cells[rowNumber, 2].Value?.ToString()!.Trim()!] =
                        worksheet.Cells[rowNumber, columnNumber].Value?.ToString()!.Trim()!;
                }
            }

            return returnDictionary;
        }
        #endregion

        #region METHODS: formatting console messages
        /// <summary>
        /// Standardized method to print messages of varying categories to the console.
        /// </summary>
        /// <param name="messageContents">The contents of the message to be written to the console.</param>
        /// <param name="messageCategory">
        /// The name of the message category. Options include:<br/>
        /// 1. 'ERROR', 'FAILURE', 'FAIL'<br/>
        /// 2. 'SUCCESS', 'PASS'<br/>
        /// 3. 'STATUS'<br/>
        /// 4. 'NEWSECTION'<br/>
        /// 5. (no category)
        /// </param>
        /// <param name="newLineForSection">
        /// A boolean input that determines whether to space a new section with the characters '---'.<br/>
        /// (Note: only applicable if messageCateogry = "NEWSECTION")
        /// </param>
        public static void ConsoleMessage(string messageContents, string messageCategory = "", bool newLineForSection = true)
        {
            messageCategory = messageCategory.ToUpper().Trim();

            if ((messageCategory == "ERROR") || (messageCategory == "FAILURE") || (messageCategory == "FAIL"))
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(messageCategory);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else if ((messageCategory == "SUCCESS") || (messageCategory == "PASS"))
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(messageCategory);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else if (messageCategory == "STATUS")
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.Write(messageCategory);
            }
            else if (messageCategory == "NEWSECTION")
            {
                if (newLineForSection)
                {
                    Console.Write($"---\n[{DateTime.Now.ToString("HH:mm:ss")}] ");
                }
                else
                {
                    Console.Write($"[{DateTime.Now.ToString("HH:mm:ss")}] ");
                }
            }
            else
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + "  ";
                Console.Write(messageCategory);
            }

            messageContents = WrapText(messageContents, 11, consoleCharLengthLimit);
            Console.WriteLine(messageContents);
        }

        /// <summary>
        /// Modify the input string to wrap the text to the next line after a certain length.<br/>
        /// The input string is seperated per word and then each line is incrementally added to per word.<br/>
        /// Start a new line when the character count of a line exceeds the specified line limit.
        /// </summary>
        /// <param name="inputString">The input string to be wrapped.</param>
        /// <param name="indentLength">An integer that defines the length of the characters in the indent starting each new line.</param>
        /// <param name="lineLimit">An integer that defines the maximum number of characters per line before a new line is created.</param>
        /// <returns>A modified string that wraps to the next line after a specified length of characters.</returns>
        private static string WrapText(string inputString, int indentLength, int lineLimit)
        {
            // Variables containing formatting information:
            StringBuilder newSentence = new StringBuilder(); // The properly formatted string to be returned.
            string[] words = inputString.Split(' ');         // An array where each element contains each word in an input string. 
            string indent = new string(' ', indentLength);   // An empty string to be used for indenting.
            string line = "";                                // The variable that will be modified and appended to the returned StringBuilder for each line.

            // Variables informing formatting logic:
            bool newLongWord = true;
            int numberOfNewLines = 0;
            int numberOfSplitWords = 0;
            int indentedLineLimit = lineLimit - indentLength;

            // Cycle through each word in the input string.
            foreach (string word in words)
            {
                // The word (short or long) has any excess spaces removed. 
                string trimmedWord = word.Trim();
                //Console.WriteLine("trimmedWord: " + trimmedWord);

                // Required for "Long Word Splitting" Logic: This variable is used to wrap long words at the indentLength specified with indenting.
                int partLengthLimit = lineLimit - (indentLength + line.Length);

                // Required for "Long Word Splitting" Logic: The # of long words determine how a long word component is added to the console.
                // Long words for this method are defined as words that are above the character number of line limit minus indent length.
                if (trimmedWord.Length >= partLengthLimit)
                    numberOfSplitWords++;

                // "Long Word Splitting" Logic
                // If the word is longer than the line limit # of characters, split it & wrap to the next line keeping indents.
                while ((trimmedWord.Length > partLengthLimit) && (trimmedWord.Length > 15))
                {
                    string part = trimmedWord.Substring(0, partLengthLimit); // A peice of the long word to add to the existing line. 
                    trimmedWord = trimmedWord.Substring(partLengthLimit);    // The long word part is removed from trimmedWord.

                    // Long Word Scenario 1: This should only ever run once the first time a long word goes through the while loop.
                    if (((numberOfSplitWords == 1) || (numberOfNewLines == 0)) && (newLongWord))
                    {
                        newSentence.AppendLine(line + part);         // Add line & part to return string. No indent b/c either the long word starts the message
                                                                     // or because the long word part gets added to the current line that already has words.
                        line = "";                                   // Reset the line string.
                        numberOfNewLines++;                          // Count up for number of new lines.
                        newLongWord = false;                         // Lock this if statement (Scenario 1) from being run again.
                        partLengthLimit = indentedLineLimit;
                    }
                    // Long Word Scenario 2: All other subsequent lines with long words (or long word components) need to be indented.
                    else
                    {
                        newSentence.AppendLine(indent + line + part);  // Add indented current line with part. (line could be 0 chars if part is long enough)
                        line = "";                                     // Reset the line string.
                        numberOfNewLines++;                            // Count up for number of new lines.
                        partLengthLimit = indentedLineLimit;
                    }
                }

                // Required for "Long Word Splitting" Logic: Determines how a long word component is added to the console.
                newLongWord = true;

                // "Adding Line" Logic
                // Check if the current line plus the next word (or the remaining part of a long word) exceeds the line limit (accounting for indenting).
                if ((line + trimmedWord).Length > indentedLineLimit)
                {
                    // Line Scenario 1: If not the first line, add indented current line to return string. 
                    if (numberOfNewLines > 0)
                    {
                        newSentence.AppendLine(indent + line.TrimEnd());
                    }
                    // Line Scenario 2: If the first line, add the current line without indents to return string.
                    else
                    {
                        newSentence.AppendLine(line.TrimEnd());
                    }
                    line = "";           // Reset the line string.
                    numberOfNewLines++;  // Count up for number of new lines.
                }

                // Add the word (or the remaining part of a long word) to the current line.
                line += trimmedWord + " ";
            }

            // Same as "Adding Line" Logic where the line contents are the remaining input string contents under the line limit. 
            if (line.Length > 0)
            {
                if (numberOfNewLines > 0)
                    newSentence.Append(indent + line.TrimEnd());
                else
                    newSentence.Append(line.TrimEnd());
            }

            return newSentence.ToString();
        }
        #endregion

        #region METHODS: get/set basic data type tags
        /// <summary>
        /// Asynchronously get the tag information of a Studio 5000 Logix Designer tag.<br/>
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="type">The data type of the tag whose value will be returned.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints updates to the console.</param>
        /// <returns>A Task that results in a S5kTag structure containing tag information (Name, Online Value, Offline Value, XPath).</returns>
        private static async Task<S5kAtomicTag> GetTagValue_Async(string XPath, DataType type, LogixProject project, bool printout = false)
        {
            S5kAtomicTag tag = new();
            string tagName = GetNameFromXPath(XPath)!;
            tag.XPath = XPath;
            tag.Name = tagName;

            try
            {
                if (type == DataType.BOOL)
                {
                    var tagValue_online = await project.GetTagValueBOOLAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueBOOLAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";

                }
                else if (type == DataType.SINT)
                {
                    var tagValue_online = await project.GetTagValueSINTAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueSINTAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";
                }
                else if (type == DataType.INT)
                {
                    var tagValue_online = await project.GetTagValueINTAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueINTAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";
                }
                else if (type == DataType.DINT)
                {
                    var tagValue_online = await project.GetTagValueDINTAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueDINTAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";
                }
                else if (type == DataType.LINT)
                {
                    var tagValue_online = await project.GetTagValueLINTAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueLINTAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";
                }
                else if (type == DataType.REAL)
                {
                    var tagValue_online = await project.GetTagValueREALAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueREALAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = $"{tagValue_offline}";
                }
                else if (type == DataType.STRING)
                {
                    var tagValue_online = await project.GetTagValueSTRINGAsync(XPath, OperationMode.Online);
                    tag.OnlineVal = (tagValue_online == "") ? "<empty_string>" : $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueSTRINGAsync(XPath, OperationMode.Offline);
                    tag.OfflineVal = (tagValue_offline == "") ? "<empty_string>" : $"{tagValue_offline}";
                }
                else
                {
                    ConsoleMessage($"Data type '{type}' not supported.", "ERROR");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Could not get tag '{tagName}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            if (printout)
            {
                string online_message = $"online value: {tag.OnlineVal}";
                string offline_message = $"offline value: {tag.OfflineVal}";
                ConsoleMessage($"{tagName.PadRight(40, ' ')}{online_message.PadRight(35, ' ')}{offline_message.PadRight(35, ' ')}", "SUCCESS");
            }

            return tag;
        }

        /// <summary>
        /// Asynchronously toggle the value of a boolean tag in Studio 5000 Logix Designer.
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="toggleOnToOff">If True, toggle the tag on to off. If False, toggle the tag off to on.</param>
        /// <param name="mode">Specify online or offline operation.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints updates to the console.</param>
        /// <returns>A Task that asynchronously toggles a specified BOOL tag.</returns>
        private static async Task ToggleBOOLTagValue_Async(string XPath, bool toggleOnToOff, OperationMode mode, LogixProject project, bool printout = false)
        {
            string tagName = GetNameFromXPath(XPath)!;
            if (toggleOnToOff)
            {
                await SetTagValue_Async(XPath, "true", mode, DataType.BOOL, project, printout);
                await SetTagValue_Async(XPath, "false", mode, DataType.BOOL, project, printout);
            }
            else
            {
                await SetTagValue_Async(XPath, "false", mode, DataType.BOOL, project, printout);
                await SetTagValue_Async(XPath, "true", mode, DataType.BOOL, project, printout);
            }
            if (printout)
                ConsoleMessage($"Done toggling '{tagName}'.", "STATUS");
        }

        /// <summary>
        /// Asynchronously set either the online or offline value of a basic data type tag.<br/>
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="newTagValue">The value of the tag that will be set.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="type">The data type of the tag whose value will be set.</param>

        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        /// <returns>A Task that will set the online or offline value of a basic data type tag.</returns>
        private static async Task SetTagValue_Async(string XPath, string newTagValue, OperationMode mode, DataType type, LogixProject project,
            bool printout = false)
        {
            string tagName = GetNameFromXPath(XPath)!;
            string oldTagValue = "";
            string newTagValueCheck = "";
            S5kAtomicTag oldTag = await GetTagValue_Async(XPath, type, project);

            try
            {
                if (type == DataType.BOOL)
                    await project.SetTagValueBOOLAsync(XPath, mode, bool.Parse(newTagValue));
                else if (type == DataType.SINT)
                    await project.SetTagValueSINTAsync(XPath, mode, sbyte.Parse(newTagValue));
                else if (type == DataType.INT)
                    await project.SetTagValueINTAsync(XPath, mode, short.Parse(newTagValue));
                else if (type == DataType.DINT)
                    await project.SetTagValueDINTAsync(XPath, mode, int.Parse(newTagValue));
                else if (type == DataType.LINT)
                    await project.SetTagValueLINTAsync(XPath, mode, long.Parse(newTagValue));
                else if (type == DataType.REAL)
                    await project.SetTagValueREALAsync(XPath, mode, float.Parse(newTagValue));
                else if (type == DataType.STRING)
                    await project.SetTagValueSTRINGAsync(XPath, mode, newTagValue);
                else
                    ConsoleMessage($"Data type '{type}' not supported.", "ERROR");
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to set tag value.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project", "ERROR");
                Console.WriteLine(e.Message);
            }

            S5kAtomicTag newTag = await GetTagValue_Async(XPath, type, project);

            if (mode == OperationMode.Online)
            {
                oldTagValue = oldTag.OnlineVal!;
                newTagValueCheck = newTag.OnlineVal!;
            }
            else if (mode == OperationMode.Offline)
            {
                oldTagValue = oldTag.OfflineVal!;
                newTagValueCheck = newTag.OfflineVal!;
            }


            if (newTag.OnlineVal!.ToUpper() != newTagValue.ToUpper())
                throw new Exception($"Tried to change '{tagName}' value to '{newTagValue}' but was '{newTagValueCheck}'.");


            if (printout)
            {
                if ((newTagValueCheck == "1") && (type == DataType.BOOL)) { newTagValueCheck = "True"; }
                if ((newTagValueCheck == "0") && (type == DataType.BOOL)) { newTagValueCheck = "False"; }

                string outputMessage = $"{oldTag.Name,-40} {oldTagValue,20} -> {newTagValueCheck,-20}";
                ConsoleMessage(outputMessage);
            }
        }
        #endregion

        #region METHODS: get/set AOI tags (complex data type)
        /// <summary>
        /// Asynchronously set (and verify the change of) multiple AOI parameter values at the same time.
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="newParameterValues">A dictionary where the keys are parameter names and where the values are the new parameter values.</param>
        /// <param name="AOIParameters">An array of the AOIParameter structure that contains required parameter name, usage, and parsing data info.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        /// <returns>A Task that will set multiple AOI parameters at the same time.</returns>
        /// <exception cref="Exception">Studio 5000 Logix Designer exceptions.</exception>
        private static async Task SetMultipleAOIParamVals_Async(string XPath, Dictionary<string, string> newParameterValues,
            AOIParameter[] AOIParameters, OperationMode mode, LogixProject project, bool printOut = false)
        {
            ByteString oldByteString = await project.GetTagValueAsync(XPath, mode, DataType.BYTE_ARRAY);
            byte[] modifiedByteArray = oldByteString.ToByteArray();

            string oldParameterValue;
            string newParameterValue;
            int numberOfInputs = 0;

            // Rotate through all the AOI parameters.
            for (int i = 0; i < AOIParameters.Length; i++)
            {
                if (GetAOIParameterComponentValue(AOIParameters[i].Name!, "Usage", AOIParameters) == "Input")
                {
                    numberOfInputs++;
                }
            }

            // Rotate through all the AOI parameters.
            for (int i = 0; i < AOIParameters.Length; i++)
            {
                // While rotating through AOI parameters, only make changes if they are not an output parameter.
                if (GetAOIParameterComponentValue(AOIParameters[i].Name!, "Usage", AOIParameters) == "Input")
                {
                    DataType dataType = GetDataType(AOIParameters[i].DataType!);
                    int bytePosition = AOIParameters[i].BytePosition;

                    oldParameterValue = AOIParameters[i].Value!;
                    newParameterValue = newParameterValues[AOIParameters[i].Name!];

                    // Update the value of the current input parameter in the byte string (currently in array format) using the data type & byte position
                    // information from AOIParameters.
                    if (dataType == DataType.BOOL)
                    {
                        byte[] bools_byteArray = new byte[4];
                        Array.ConstrainedCopy(modifiedByteArray, bytePosition, bools_byteArray, 0, 4);
                        var bitArray = new BitArray(bools_byteArray);

                        int boolPosition = AOIParameters[i].BoolPosition;
                        bool bool_newTagValue = newParameterValue == "1";
                        bitArray[boolPosition] = bool_newTagValue;
                        bitArray.CopyTo(bools_byteArray, 0);


                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = bools_byteArray[j];
                    }
                    else if (dataType == DataType.SINT)
                    {
                        string sint_string = Convert.ToString(long.Parse(newParameterValue), 2);
                        modifiedByteArray[bytePosition] = Convert.ToByte(sint_string, 2);
                    }
                    else if (dataType == DataType.INT)
                    {
                        byte[] int_byteArray = BitConverter.GetBytes(int.Parse(newParameterValue));
                        for (int j = 0; j < 2; ++j)
                            modifiedByteArray[j + bytePosition] = int_byteArray[j];
                    }
                    else if (dataType == DataType.DINT)
                    {
                        byte[] dint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = dint_byteArray[j];
                    }
                    else if (dataType == DataType.LINT)
                    {
                        byte[] lint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int j = 0; j < 8; ++j)
                            modifiedByteArray[j + bytePosition] = lint_byteArray[j];
                    }
                    else if (dataType == DataType.REAL)
                    {
                        byte[] real_byteArray = BitConverter.GetBytes(float.Parse(newParameterValue));
                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = real_byteArray[j];
                    }
                    else
                    {
                        ConsoleMessage($"Data type '{dataType}' not supported by 'SetMultipleAOIParamVals_Async' method.", "ERROR");
                    }

                    // If specified using the method input, print the current parameter changed to the console.
                    if (printOut)
                    {
                        string setParamIntro = $"{AOIParameters[i].Name} value:".PadRight(40, ' ');

                        // Write the first formatting component of the current parameter to the console.
                        if (i == 0)
                        {
                            setParamIntro = "┌── " + setParamIntro;
                        }
                        else if (i < numberOfInputs)
                        {
                            setParamIntro = "├── " + setParamIntro;
                        }
                        else
                        {
                            setParamIntro = "└── " + setParamIntro;
                        }

                        ConsoleMessage($"{setParamIntro} {oldParameterValue,20} -> {newParameterValue,-20}");
                    }
                }
            }

            // Push the newly modified byte string (containing the AOI tag information) to Studio 5000 Logix Designer. 
            try
            {
                await project.SetTagValueAsync(XPath, mode, modifiedByteArray, DataType.BYTE_ARRAY);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to set tag values.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Verify that the tag change was actually implemented.
            try
            {
                ByteString newByteString = await project.GetTagValueAsync(XPath, mode, DataType.BYTE_ARRAY);
                AOIParameter[] newAOIParameters = GetAOIParameterValues(AOIParameters, newByteString);

                foreach (var kvp in newParameterValues)
                {
                    foreach (var param in newAOIParameters)
                    {
                        if ((kvp.Key == param.Name) && (param.Usage == "Input"))
                        {
                            if (kvp.Value != param.Value)
                            {
                                throw new Exception("Method 'SetMultipleAOIParamVals_Async' did not properly set tag values.");
                            }
                        }
                    }
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to get new tag values to verify tag change.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Save the project.
            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project.", "ERROR");
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Parse any AOI ByteString into its relevant parameter values.<br/>
        /// To parse the input ByteString properly, the following conditions must be provided in the input AOIParameter array:<br/>
        /// 1. The data types of all the AOI parameters must already be included.<br/>
        /// 2. The AOI parameter data types must be ordered how they were added during AOI creation (as shown in the AOI instruction window in S5k).
        /// </summary>
        /// <param name="AOIParameters">An array of the AOIParameter structure that contains required parameter name, usage, and parsing data info.</param>
        /// <param name="aoiByteString">The byte string from which to get values for the AOIParameters input.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        /// <returns>An updated instance of the AOIParameter structure with parameter values from the input ByteString.</returns>
        public static AOIParameter[] GetAOIParameterValues(AOIParameter[] AOIParameters, ByteString aoiByteString, bool printOut = false)
        {
            // Split ByteSTring into its individual bytes.
            byte[] inputByteArray = aoiByteString.ToByteArray();

            // To properly parse the byte array, the below variables track and/or create the logic determining which bytes to get information from.
            int byteStartPosition = 0; // Increases by different amounts based on the data type of the parameter (ex. DINTs = 4 bytes, LINTs = 8 bytes).
            int boolStartPosition = 0; // Store the boolean's byte location because they are stored together.
            int boolCount = 0;         // Tracks the # of booleans. This helps determine when to update the boolStartPosition.

            // Rotate through all the AOI parameters.
            for (int i = 0; i < AOIParameters.Length; i++)
            {
                // Get the data type of the current element (AOI parameter) of the AOIParameters array.
                string paramDataType = AOIParameters[i].DataType!;

                // Based on the current AOI parameter's data type, convert a specific number of bytes into the parameter value. Then, update the byte trackers.
                if (paramDataType == "BOOL")
                {
                    // Update the "boolean host member" location of the input byte array that is being checked every 32 booleans.
                    if (((boolCount % 32 == 1) && (boolCount > 1)) || (boolCount == 0))
                    {
                        boolStartPosition = byteStartPosition;
                        byteStartPosition += 4;
                    }

                    // Booleans are stored in clusters of 4 bytes.
                    byte[] bools_bytearray = new byte[4];
                    Array.ConstrainedCopy(inputByteArray, boolStartPosition, bools_bytearray, 0, 4);

                    // Reverse the elements of the byte array & put them all in one string. This helps to get a specific boolean's value later.
                    StringBuilder sb = new StringBuilder();
                    for (int j = bools_bytearray.Length - 1; j >= 0; j--)
                        sb.Append(Convert.ToString(bools_bytearray[j], 2).PadLeft(8, '0'));
                    string bools_string = sb.ToString();

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = bools_string[31 - boolCount].ToString();
                    AOIParameters[i].BytePosition = boolStartPosition;
                    AOIParameters[i].BoolPosition = boolCount;

                    boolCount++;
                }
                else if (paramDataType == "SINT")
                {
                    // Single integers are 1 byte in length.
                    byte[] sint_bytearray = new byte[1];
                    Array.ConstrainedCopy(inputByteArray, byteStartPosition, sint_bytearray, 0, 1);
                    string sint_string = Convert.ToString(unchecked((sbyte)sint_bytearray[0]));

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = sint_string;
                    AOIParameters[i].BytePosition = byteStartPosition;
                    byteStartPosition += 1;
                }
                else if (paramDataType == "INT")
                {
                    // Integers are not stored at "odd valued" byte locations. If odd, make the tracker even. Consider the skipped, empty byte a "buffer byte".
                    if ((byteStartPosition % 2) > 0)
                        byteStartPosition += 1;

                    // Integers are 1 byte in length.
                    byte[] int_bytearray = new byte[2];
                    Array.ConstrainedCopy(inputByteArray, byteStartPosition, int_bytearray, 0, 2);
                    string int_string = Convert.ToString(BitConverter.ToInt16(int_bytearray));

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = int_string;
                    AOIParameters[i].BytePosition = byteStartPosition;
                    byteStartPosition += 2;
                }
                else if (paramDataType == "DINT")
                {
                    if ((byteStartPosition % 4) > 0)
                        byteStartPosition += 4 - (byteStartPosition % 4);

                    byte[] dint_bytearray = new byte[4];
                    Array.ConstrainedCopy(inputByteArray, byteStartPosition, dint_bytearray, 0, 4);
                    string dint_string = Convert.ToString(BitConverter.ToInt32(dint_bytearray));

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = dint_string;
                    AOIParameters[i].BytePosition = byteStartPosition;
                    byteStartPosition += 4;
                }
                else if (paramDataType == "LINT")
                {
                    if ((byteStartPosition % 8) > 0)
                        byteStartPosition += 8 - (byteStartPosition % 8);

                    byte[] lint_bytearray = new byte[8];
                    Array.ConstrainedCopy(inputByteArray, byteStartPosition, lint_bytearray, 0, 8);
                    string lint_string = Convert.ToString(BitConverter.ToInt64(lint_bytearray));

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = lint_string;
                    AOIParameters[i].BytePosition = byteStartPosition;
                    byteStartPosition += 8;
                }
                else if (paramDataType == "REAL")
                {
                    if ((byteStartPosition % 4) > 0)
                        byteStartPosition += 4 - (byteStartPosition % 4);

                    byte[] real_bytearray = new byte[4];
                    Array.ConstrainedCopy(inputByteArray, byteStartPosition, real_bytearray, 0, 4);
                    string real_string = Convert.ToString(BitConverter.ToSingle(real_bytearray));

                    // Update parameter value & tracker information to the "final result array". 
                    AOIParameters[i].Value = real_string;
                    AOIParameters[i].BytePosition = byteStartPosition;
                    byteStartPosition += 4;
                }
                else
                {
                    ConsoleMessage($"The GetAOIParameterValues method cannot handle process the data type '{paramDataType}'.", "STATUS");
                    AOIParameters[i].BytePosition = byteStartPosition;
                }
            }

            // Print all the AOI parameter structure components to the console.
            if (printOut)
            {
                PrintAOIParameters(AOIParameters, true);
            }

            return AOIParameters;
        }

        /// <summary>
        /// Print all the AOI information within an AOIParameter structure array to the console.
        /// </summary>
        /// <param name="AOIParameters">The AOIParameter structure array to be printed.</param>
        /// <param name="printAll">If True, print the internally used structure subcomponents, BytePosition and BoolPosition.</param>
        private static void PrintAOIParameters(AOIParameter[]? AOIParameters, bool printAll)
        {
            // Console formatting: Get the max character length of the below 4 AOIParameter structure subcomponents within the input array.
            int[] AOIParameterLimits = new int[5];
            AOIParameterLimits[3] = 5;
            AOIParameterLimits[4] = 20;
            for (int i = 0; i < AOIParameters!.Length; i++)
            {
                if (AOIParameters[i].Value == null)
                    AOIParameters[i].Value = "";
                if (AOIParameters[i].XPath == null)
                    AOIParameters[i].XPath = "";

                if (AOIParameters[i].Name!.Length > AOIParameterLimits[0])
                    AOIParameterLimits[0] = AOIParameters[i].Name!.Length;
                if (AOIParameters[i].DataType!.Length > AOIParameterLimits[1])
                    AOIParameterLimits[1] = AOIParameters[i].DataType!.Length;
                if (AOIParameters[i].Usage!.Length > AOIParameterLimits[2])
                    AOIParameterLimits[2] = AOIParameters[i].Usage!.Length;
                if (AOIParameters[i].Value!.Length > AOIParameterLimits[3])
                    AOIParameterLimits[3] = AOIParameters[i].Value!.Length;
                if (AOIParameters[i].XPath!.Length > AOIParameterLimits[4])
                    AOIParameterLimits[4] = AOIParameters[i].XPath!.Length;
            }

            for (int i = 0; i < AOIParameters.Length; i++)
            {
                // Write the first formatting component of a new parameter to the console.
                if (i == 0)
                {
                    Console.Write(" ┌── ");
                }
                else if (i < AOIParameters.Length - 1)
                {
                    Console.Write(" ├── ");
                }
                else
                {
                    Console.Write(" └── ");
                }

                // Add the parameter formatted information to the current line.
                if (printAll)
                {
                    Console.WriteLine($"Name: {AOIParameters[i].Name!.PadRight(AOIParameterLimits[0], ' ')} | " +
                        $"Data Type: {AOIParameters[i].DataType!.PadRight(AOIParameterLimits[1], ' ')} | " +
                        $"Usage: {AOIParameters[i].Usage!.PadRight(AOIParameterLimits[2], ' ')} | Required: {AOIParameters[i].Required,-5} | " +
                        $"Visible: {AOIParameters[i].Visible,-5} |  Value: {AOIParameters[i].Value!.PadRight(AOIParameterLimits[3], ' ')} | " +
                        $"Byte Position: {AOIParameters[i].BytePosition,-3} | Bool Position: {AOIParameters[i].BoolPosition} | " +
                        $"XPath: {AOIParameters[i].XPath!.PadRight(AOIParameterLimits[4], ' ')}");
                }
                else
                {
                    Console.WriteLine($"Name: {AOIParameters[i].Name!.PadRight(AOIParameterLimits[0], ' ')} | " +
                        $"Data Type: {AOIParameters[i].DataType!.PadRight(AOIParameterLimits[1], ' ')} | " +
                        $"Usage: {AOIParameters[i].Usage!.PadRight(AOIParameterLimits[2], ' ')} |  " +
                        $"Value: {AOIParameters[i].Value!.PadRight(AOIParameterLimits[3], ' ')}");
                }
            }
        }

        /// <summary>
        /// Get a specific subcomponent value from a specific element of the AOIParameter structure array. 
        /// </summary>
        /// <param name="parameterName">The name of the AOI parameter that determines which element of the AOIParameter array to search within.</param>
        /// <param name="AOIParameterField">The target field (structure subcomponent name) to get the value of.<br/>
        /// (Name, DataType, Usage, Required, Visible, Value, BytePosition, BoolPosition)</param>
        /// <param name="AOIParameters">An array of the AOIParameter structure that contains the parameter characteristics and parsing data info.</param>
        /// <returns>The value of the target field (structure subcomponent name) as a string.</returns>
        public static string GetAOIParameterComponentValue(string parameterName, string AOIParameterField, AOIParameter[] AOIParameters)
        {
            AOIParameterField = AOIParameterField.Trim().ToUpper();
            string returnString = "";
            for (int i = 0; i < AOIParameters.Length; i++)
            {
                if (AOIParameters[i].Name == parameterName)
                {
                    if (AOIParameterField == "NAME")
                    {
                        returnString = AOIParameters[i].Name!;
                    }
                    if (AOIParameterField == "DATATYPE")
                    {
                        returnString = AOIParameters[i].DataType!;
                    }
                    if (AOIParameterField == "USAGE")
                    {
                        returnString = AOIParameters[i].Usage!;
                    }
                    if (AOIParameterField == "REQUIRED")
                    {
                        returnString = AOIParameters[i].Required.ToString()!;
                    }
                    if (AOIParameterField == "VISIBLE")
                    {
                        returnString = AOIParameters[i].Visible.ToString()!;
                    }
                    if (AOIParameterField == "VALUE")
                    {
                        returnString = AOIParameters[i].Value!;
                    }
                    if (AOIParameterField == "BYTEPOSITION")
                    {
                        returnString = AOIParameters[i].BytePosition.ToString();
                    }
                    if (AOIParameterField == "BOOLPOSITION")
                    {
                        returnString = AOIParameters[i].BoolPosition.ToString();
                    }
                    if (AOIParameterField == "XPATH")
                    {
                        returnString = AOIParameters[i].XPath!;
                    }
                }
            }
            return returnString;
        }

        /// <summary>
        /// Asynchronously set (and verify the change of) a single value of an AOI tag in Studio 5000 Logix Designer.
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        ///</param>
        /// <param name="newParameterValue">The new value of the AOI parameter as a string.</param>
        /// <param name="parameterName">The name of the parameter that will have its value changed.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="AOIParameters">An array of the AOIParameter structure that contains required parameter name, usage, and parsing data info.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printOut">A boolean that, if True, prints updates to the console.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Studio 5000 Logix Designer exceptions.</exception>
        private static async Task SetSingleAOIParamValue_Async(string XPath, string newParameterValue, string parameterName, OperationMode mode,
            AOIParameter[] AOIParameters, LogixProject project, bool printOut = false)
        {
            ByteString oldByteString = await project.GetTagValueAsync(XPath, mode, DataType.BYTE_ARRAY);
            AOIParameter[] oldAOIParameters = GetAOIParameterValues(AOIParameters, oldByteString);

            byte[] modifiedByteArray = oldByteString.ToByteArray();

            string oldParameterValue = GetAOIParameterComponentValue(parameterName, "Value", oldAOIParameters);

            // Rotate through all the AOI parameters.
            for (int i = 0; i < AOIParameters.Length; i++)
            {
                // Stop rotating through the AOI parameters when on the parameter 
                if (AOIParameters[i].Name == parameterName)
                {
                    DataType dataType = GetDataType(AOIParameters[i].DataType!);
                    int bytePosition = AOIParameters[i].BytePosition;

                    if (dataType == DataType.BOOL)
                    {
                        byte[] bools_byteArray = new byte[4];
                        Array.ConstrainedCopy(modifiedByteArray, bytePosition, bools_byteArray, 0, 4);
                        var bitArray = new BitArray(bools_byteArray);

                        int boolPosition = AOIParameters[i].BoolPosition;
                        bool bool_newTagValue = bool.Parse(newParameterValue);
                        newParameterValue = (newParameterValue == "True") ? "1" : "0";
                        bitArray[boolPosition] = bool_newTagValue;
                        bitArray.CopyTo(bools_byteArray, 0);


                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = bools_byteArray[j];
                    }
                    else if (dataType == DataType.SINT)
                    {
                        string sint_string = Convert.ToString(long.Parse(newParameterValue), 2);
                        sint_string = sint_string.Substring(sint_string.Length - 8);
                        modifiedByteArray[bytePosition] = Convert.ToByte(sint_string, 2);
                    }
                    else if (dataType == DataType.INT)
                    {
                        byte[] int_byteArray = BitConverter.GetBytes(int.Parse(newParameterValue));
                        for (int j = 0; j < 2; ++j)
                            modifiedByteArray[j + bytePosition] = int_byteArray[j];
                    }
                    else if (dataType == DataType.DINT)
                    {
                        byte[] dint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = dint_byteArray[j];
                    }
                    else if (dataType == DataType.LINT)
                    {
                        byte[] lint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int j = 0; j < 8; ++j)
                            modifiedByteArray[j + bytePosition] = lint_byteArray[j];
                    }
                    else if (dataType == DataType.REAL)
                    {
                        byte[] real_byteArray = BitConverter.GetBytes(float.Parse(newParameterValue));
                        for (int j = 0; j < 4; ++j)
                            modifiedByteArray[j + bytePosition] = real_byteArray[j];
                    }
                    else
                    {
                        ConsoleMessage($"Data type '{dataType}' not supported by 'SetSingleAOIParamValue_Async' method.", "ERROR");
                    }
                }
            }

            // Push the newly modified byte string (containing the AOI tag information) to Studio 5000 Logix Designer. 
            try
            {
                await project.SetTagValueAsync(XPath, mode, modifiedByteArray, DataType.BYTE_ARRAY);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to set tag value.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Verify that the tag value change was actually implemented.
            try
            {
                ByteString newByteString = await project.GetTagValueAsync(XPath, mode, DataType.BYTE_ARRAY);
                AOIParameter[] updatedAOIParameters = GetAOIParameterValues(AOIParameters, newByteString);
                string updatedParameterValue = GetAOIParameterComponentValue(parameterName, "Value", updatedAOIParameters);

                if (updatedAOIParameters != oldAOIParameters)
                {
                    throw new Exception("SetSingleAOIParamValue_Async method did not properly set tag value.");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to get new tag values to verify tag change.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Save the project.
            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project.", "ERROR");
                Console.WriteLine(e.Message);
            }

            if (printOut)
            {
                string setParamIntro = $"{parameterName} value:".PadRight(40, ' ');
                ConsoleMessage($"{setParamIntro} {oldParameterValue,20} -> {newParameterValue,-20}");
            }
        }
        #endregion

        #region METHODS: read/change controller mode & download
        /// <summary>
        /// Asynchronously get the current controller mode (FAULTED, PROGRAM, RUN, or TEST).
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>A Task that returns a string of the current controller mode.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the returned controller mode is not FAULTED, PROGRAM, RUN, or TEST.</exception>
        private static async Task<string> ReadControllerMode_Async(string commPath, LogixProject project)
        {
            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set commpath to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                LogixProject.ControllerMode result = await project.ReadControllerModeAsync();
                switch (result)
                {
                    case LogixProject.ControllerMode.Faulted:
                        return "FAULTED";
                    case LogixProject.ControllerMode.Program:
                        return "PROGRAM";
                    case LogixProject.ControllerMode.Run:
                        return "RUN";
                    case LogixProject.ControllerMode.Test:
                        return "TEST";
                    default:
                        throw new ArgumentOutOfRangeException("Controller mode is unrecognized.");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to read controller mode.", "ERROR");
                Console.WriteLine(e.Message);
            }

            return "";
        }

        /// <summary>
        /// Asynchronously change the controller mode to either Program, Run, or Test mode.
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="mode">The controller mode to switch to.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>A Task that changes the controller mode.</returns>
        private static async Task ChangeControllerMode_Async(string commPath, string mode, LogixProject project)
        {
            mode = mode.ToUpper().Trim();

            var requestedControllerMode = default(LogixProject.RequestedControllerMode);
            if (mode == "PROGRAM")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Program;
            }
            else if (mode == "RUN")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Run;
            }
            else if (mode == "TEST")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Test;
            }
            else
            {
                ConsoleMessage($"Mode '{mode}' is not supported.", "ERROR");
            }

            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set communication path to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.ChangeControllerModeAsync(requestedControllerMode);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set mode. Requested mode was '{mode}'.", "ERROR");
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Asynchronously download to the specified controller.
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>An Task that downloads to the specified controller.</returns>
        private static async Task DownloadProject_Async(string commPath, LogixProject project)
        {
            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set communication path to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                LogixProject.ControllerMode controllerMode = await project.ReadControllerModeAsync();
                if (controllerMode != LogixProject.ControllerMode.Program)
                {
                    ConsoleMessage($"Controller mode is {controllerMode}. Downloading is possible only if the controller is in 'Program' mode.", "ERROR");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to read ControllerMode.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.DownloadAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to download.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Download modifies the project.
            // Without saving, if used file will be opened again, commands which need correlation
            // between program in the controller and opened project like LoadImageFromSDCard or StoreImageOnSDCard
            // may not be able to succeed because project in the controller won't match opened project.
            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project.", "ERROR");
                Console.WriteLine(e.Message);
            }
        }
        #endregion

        #region METHODS: TEST & helper methods
        /// <summary>
        /// A test to compare the expected and actual values of a tag.
        /// </summary>
        /// <param name="tagName">The name of the tag to be tested.</param>
        /// <param name="expectedValue">The expected value of the tag under test.</param>
        /// <param name="actualValue">The actual value of the tag under test.</param>
        /// <returns>Return an integer value 1 for test failure and an integer value 0 for test success.</returns>
        /// <remarks>
        /// The integer output is added to an integer that tracks the total number of failures.<br/>
        /// At the end of all testing, the overall SUCCESS/FAILURE of this CI/CD test stage is determined whether its value is greater than 0.
        /// </remarks>
        private static int TEST_CompareForExpectedValue(string tagName, string expectedValue, string actualValue, bool printOut)
        {
            if (expectedValue != actualValue)
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' NOT equal.", "FAIL");

                return 1;
            }
            else
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' EQUAL.", "PASS");

                return 0;
            }
        }

        /// <summary>
        /// Helper method for the GetInOutParamCDATA_forData method.<br/>
        /// Return the CDATA format of the specified datatype.
        /// </summary>
        /// <param name="dataType">The name of the target datatype.</param>
        /// <param name="IsArray">Boolean defaulted to false that varies what kind of CDATA the bool data type returns.</param>
        /// <returns>A string in the L5X CDATA format of a particular datatype.</returns>
        public static string? GetS5kAtomicTagCDATA_forData(string dataType, bool IsArray = false)
        {
            dataType = dataType.ToUpper();
            string boolcdataNoArray = "0";
            string boolcdataWithArray = "2#0";
            string sintcdata = "0";
            string intcdata = "0";
            string dintcdata = "0";
            string lintcdata = "0";
            string realcdata = "0.00000000e+000";
            string stringcdata = "[0,'$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00"
                          + "$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00$00']";

            if ((dataType == "BOOL") || (dataType == "BIT"))
            {
                if (IsArray)
                    return boolcdataWithArray;
                else
                    return boolcdataNoArray;
            }
            else if (dataType == "SINT")
            {
                return sintcdata;
            }
            else if (dataType == "INT")
            {
                return intcdata;
            }
            else if (dataType == "DINT")
            {
                return dintcdata;
            }
            else if (dataType == "LINT")
            {
                return lintcdata;
            }
            else if (dataType == "REAL")
            {
                return realcdata;
            }
            else if (dataType == "STRING")
            {
                return stringcdata;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Helper method for the GetInOutParamCDATA_forData method.<br/>
        /// Create a the format needed for a CDATA array.
        /// </summary>
        /// <param name="objectToRepeat">The string to be repeated.</param>
        /// <param name="repeatCount">The number of times to repeat the string.</param>
        /// <returns>A formatted string in the L5X array format.</returns>
        public static string CreateArrayCDATA(string objectToRepeat, int repeatCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append('[');

            for (int i = 0; i < repeatCount; i++)
            {
                if (i > 0)
                {
                    sb.Append(',');
                }
                sb.Append(objectToRepeat);
            }

            sb.Append(']');
            return sb.ToString();
        }

        /// <summary>
        /// Method to replace the string representation of a tag data type with the LDSDK provided DataType enumerator.
        /// </summary>
        /// <param name="dataType">The name of the data type to be returned.</param>
        /// <returns>The LDSDK provided DataType enumerator</returns>
        /// <exception cref="ArgumentException"></exception>
        private static DataType GetDataType(string dataType)
        {
            DataType type;
            switch (dataType)
            {
                case "BOOL":
                    type = DataType.BOOL;
                    break;
                case "SINT":
                    type = DataType.SINT;
                    break;
                case "INT":
                    type = DataType.INT;
                    break;
                case "DINT":
                    type = DataType.DINT;
                    break;
                case "REAL":
                    type = DataType.REAL;
                    break;
                case "LINT":
                    type = DataType.LINT;
                    break;
                default:
                    ConsoleMessage($"Data type '{dataType}' not supported.", "ERROR");
                    throw new ArgumentException();
            }
            return type;
        }

        /// <summary>
        /// Get the contents within the right-most brackets and apostrophes of the input XPath.<br/>
        /// </summary>
        /// <param name="XPath">
        /// The XPath (tag path) specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The XPath for a tag is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <returns>
        /// The contents of an XPath.<br/>
        /// Ex.) Controller/Tags/Tag[@Name='AT_ClearFault'] returns AT_ClearFault
        /// </returns>
        public static string? GetNameFromXPath(string XPath)
        {
            string pattern = @"'([^']*)'";
            MatchCollection matches = Regex.Matches(XPath, pattern);

            if (matches.Count > 0)
                return matches[matches.Count - 1].Groups[1].Value;
            else
            {
                ConsoleMessage($"Tag path '{XPath}' does not contain an [@Name=''] or [@Class=''] region.", "ERROR");
                return null;
            }
        }

        /// <summary>
        /// Create a copy of a specified file at a specific location for an AOI definition to routine conversion.
        /// </summary>
        /// <param name="sourceFilePath">The file to be copied.</param>
        /// <param name="outputFolderPath">The folder path in which the new file is to be saved.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>The file path of the copied file.</returns>
        public static string CopyFile(string sourceFilePath, string outputFolderPath, bool printOut = false)
        {
            if (!File.Exists(sourceFilePath))
                ConsoleMessage($"Source file '{sourceFilePath}' does not exist.", "ERROR");

            string targetObjectName = GetAttributeValue(sourceFilePath, "AddOnInstructionDefinition", "Name", printOut)!;

            // Get the directory and file name from the source file path.
            string extension = Path.GetExtension(sourceFilePath);

            // Construct the new file path for the copied file.
            string newFileName = $"{currentDateTime}_{targetObjectName}_AOIRoutine{extension}";
            string newFilePath = Path.Combine(outputFolderPath, newFileName);

            File.Copy(sourceFilePath, newFilePath, overwrite: true);

            return newFilePath;
        }
        #endregion
    }
}