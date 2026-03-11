using LogixEcho;
using RockwellAutomation.LogixDesigner;
using System.Text;

namespace SimpleExample
{
    public class SimpleExampleClass
    {
        static async Task Main()
        {
            #region SET UP TESTING: create an emulated chassis & controller for our application & go online with the specified simple test case application.
            // Record when testing started.
            DateTime testStartTime = DateTime.Now;

            // Print unit test banner to the console.
            Console.WriteLine("\n  ".PadRight(120 - 2, '='));
            Console.WriteLine("".PadRight(120, '='));
            string bannerContents = "UNIT TESTING | " + testStartTime + " " + TimeZoneInfo.Local;
            int padding = (120 - bannerContents.Length) / 2;
            Console.WriteLine(bannerContents.PadLeft(bannerContents.Length + padding).PadRight(120));
            Console.WriteLine("".PadRight(120, '='));
            Console.WriteLine("  ".PadRight(120 - 2, '=') + "\n");

            // The file path of the simple Studio 5000 Logix Designer application to be tested.
            string acdFilePath = @"C:\Lab Files\Automated Testing\Simple\ACD_UnderTest\LDSDKdemo_simpleEx_L85E_v36.ACD";

            // Set up emulated controller (based on the specified ACD file path & unit test static variables) if one does not yet exist.
            ConsoleMessage("START setting up Factory Talk Logix Echo emulated controller...", "NEWSECTION", false);
            string commPath = LogixEchoMethods.Main(acdFilePath, "SimpleLDSDKEx_Chassis", "LDSDKdemo_simpleEx_L85E_v36").GetAwaiter().GetResult();
            ConsoleMessage($"Project communication path specified is '{commPath}'.", "STATUS");

            // Create an instance of the LogixProject class that will be used to configure/manipulate the controller.
            LogixProject logixProject = await LogixProject.OpenLogixProjectAsync(acdFilePath);

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
            #endregion

            #region COMMENCE TESTING: set & check parameters for various simple test cases.
            ConsoleMessage($"START simple testing...", "NEWSECTION");

            int failureCondition = 0;
            string box_length_XPath = "Controller/Tags/Tag[@Name='box_length']";                                     // Controller scoped (input)
            string box_width_XPath = "Controller/Tags/Tag[@Name='box_width']";                                       // Controller scoped (input)
            string area_comparison1_result_XPath = "Controller/Tags/Tag[@Name='area_comparison_result1']";             // Controller scoped (output)
            string box_area_XPath = "Controller/Programs/Program[@Name='P00_Program']/Tags/Tag[@Name='box_area']"; // Program scoped    (output)
            //****Define string for comparison 2 here****
            

            ConsoleMessage($"START test case 1/3...", "STATUS");

            // Test case 1: set inputs
            await logixProject.SetTagValueSINTAsync(box_length_XPath, LogixProject.OperationMode.Online, 3);
            await logixProject.SetTagValueDINTAsync(box_width_XPath, LogixProject.OperationMode.Online, 3);

            // Test case 1: get outputs
            long testCase1_area = await logixProject.GetTagValueLINTAsync(box_area_XPath, LogixProject.OperationMode.Online);
            bool testCase1_cmpResult1 = await logixProject.GetTagValueBOOLAsync(area_comparison1_result_XPath, LogixProject.OperationMode.Online);
            //****Add code here to read second comparison value for Test case 1****
            

            // Test case 1: verify outputs
            // if length = 3 & width = 3, then area should = 9
            failureCondition += TEST_CompareForExpectedValue("box_area", "9", testCase1_area.ToString(), true);
            // if area = 9 then area_comparison_result1 should = False (area 9 not greater than area comparison 10)
            failureCondition += TEST_CompareForExpectedValue("area_comparison_result1", "FALSE", testCase1_cmpResult1.ToString().ToUpper(), true);
            //****Add code here to test second comparison value for Test case 1****
            

            ConsoleMessage($"START test case 2/3...", "STATUS");

            // Test case 2: set inputs
            await logixProject.SetTagValueSINTAsync(box_length_XPath, LogixProject.OperationMode.Online, 4);
            await logixProject.SetTagValueDINTAsync(box_width_XPath, LogixProject.OperationMode.Online, 4);

            // Test case 2: get outputs
            long testCase2_area = await logixProject.GetTagValueLINTAsync(box_area_XPath, LogixProject.OperationMode.Online);
            bool testCase2_cmpResult1 = await logixProject.GetTagValueBOOLAsync(area_comparison1_result_XPath, LogixProject.OperationMode.Online);
            //****Add code here to read second comparison for Test case 2****
            

            // Test case 2: verify outputs
            // if length = 4 & width = 4, then area should = 16
            failureCondition += TEST_CompareForExpectedValue("box_area", "16", testCase2_area.ToString(), true);
            // if area = 9 then area_comparison_result1 should = True (area 16 greater than area comparison 10)
            failureCondition += TEST_CompareForExpectedValue("area_comparison_result1", "TRUE", testCase2_cmpResult1.ToString().ToUpper(), true);
            //****Add code here to test second comparison value for Test case 2****
            

            ConsoleMessage($"START test case 3/3...", "STATUS");

            // Test case 3: set inputs
            await logixProject.SetTagValueSINTAsync(box_length_XPath, LogixProject.OperationMode.Online, 5);
            await logixProject.SetTagValueDINTAsync(box_width_XPath, LogixProject.OperationMode.Online, 5);

            // Test case 3: get outputs
            long testCase3_area = await logixProject.GetTagValueLINTAsync(box_area_XPath, LogixProject.OperationMode.Online);
            bool testCase3_cmpResult1 = await logixProject.GetTagValueBOOLAsync(area_comparison1_result_XPath, LogixProject.OperationMode.Online);
            //****Add code here to read second comparison for Test case 3****
            

            // Test case 3: verify outputs
            // if length = 5 & width = 5, then area should = 25
            failureCondition += TEST_CompareForExpectedValue("box_area", "25", testCase3_area.ToString(), true);
            // if area = 25 then area_comparison_result1 should = True (area 25 greater than area comparison 10)
            failureCondition += TEST_CompareForExpectedValue("area_comparison_result1", "TRUE", testCase3_cmpResult1.ToString().ToUpper(), true);
            //****Add code here to test second comparison value for Test case 3****
            
            // Based on the AOI unit test result, print a final result message in red or green.
            if (failureCondition > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ConsoleMessage($"Overall Test Final Result: FAIL | {failureCondition} Issues Encountered", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                ConsoleMessage($"Overall Unit Test Final Result: PASS", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            #endregion

            #region END TEST: print final messages to the console.
            // Compute how long the test took to run. 
            DateTime testEndTime = DateTime.Now;
            TimeSpan testLength = testEndTime.Subtract(testStartTime);
            string formattedTestLength = testLength.ToString(@"hh\:mm\:ss");
            ConsoleMessage($"Testing completed in {formattedTestLength} (HH:mm:ss).", "NEWSECTION");

            await logixProject.GoOfflineAsync(); // Testing is complete. Go offline with the emulated controller.
            #endregion
        }
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

            messageContents = WrapText(messageContents, 11, 120);
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
        #endregion
    }
}