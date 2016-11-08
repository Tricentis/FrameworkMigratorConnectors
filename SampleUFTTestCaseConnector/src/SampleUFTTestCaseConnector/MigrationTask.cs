using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ICSharpCode.SharpZipLib.Zip;
using TCMigrationAPI;
using TCMigrationAPI.UFT;

namespace SampleUFTTestCaseConnector
{
    /// <summary>
    /// This class contains the frmaework specific business logic for migration.
    /// Inheriting this class from ObjectImporter class is mandatory, as it instantiates the Definition and Builder object from TCMigrationAPI in the base constructor with necessary parameters.
    /// Definition and Builder object will be available in the class and provides all utility methods necessary for migration.
    /// </summary>
    public class MigrationTask : ObjectImporter
    {
        //Derived class constructor
        public MigrationTask(ToscaObjectDefinition definition, string engine) : base(definition, engine)
        {
        }

        /// <summary>
        /// This method is the entry point of MigrationTask and consumes the file path of the archive.
        /// </summary>
        /// <param name="filePath">File path of the archive.</param>
        public override void ProcessArchive(string filePath)
        {
            //ICSharpCode.SharpZipLib library has been used to extract the archive. The library can be found in Tosca Installation Directory at '%TRICENTIS_HOME%\ToscaCommander'.
            FastZip archive = new FastZip();
            archive.ExtractZip(filePath, Definition.MigrationFolderPath, null);

            //The TestCase folder structure of the archive can be modified from here.
            string[] testCaseFiles = Directory.GetFiles(Definition.MigrationFolderPath + "Framework\\TestCase");

            //Processes all the TestCase files in the archive.
            foreach (string file in testCaseFiles)
            {
                ParseTestScriptFile(file);
            }
        }

        /// <summary>
        /// This method migrates the TestCase excel files.
        /// </summary>
        /// <param name="filePath">File path of the TestCase excel.</param>
        protected override void ParseTestScriptFile(string filePath)
        {
            //ClosedXML has been used here to read the excel files. The library can be found in Tosca Installation Directory at '%TRICENTIS_HOME%\ToscaCommander'.
            //Alternatively, Microsoft.Office.Interop.Excel library or other third-party library can also be used.
            XLWorkbook workBook = new XLWorkbook(filePath);

            foreach (IXLWorksheet sheet in workBook.Worksheets)
            {
                IXLRange usedRange = sheet.RangeUsed();
                //Builder object provides all the methods to create Tosca Business Objects. Here we create a TestCase. Also, Definition object provide the root folder IDs for import.
                int testCaseId = Builder.CreateTestCase(sheet.Name, Definition.TestCasesFolderId);
                string objectType = null;
                string browser = null, keyword = null, data = null;
                string browserName = null, pageName = null, controlName = null;
                int moduleId = 0,
                    testStepId = 0,
                    testSheetId = 0;
                bool isTestCaseTemplate = false;
                string objectrepositoryFolderPath = Definition.MigrationFolderPath + "Framework\\ObjectRepository";

                for (int row = 2; row <= usedRange.RowCount(); row++)
                {
                    Dictionary<string, string> technicalIdParam = new Dictionary<string, string>();

                    for (int column = 1; column <= usedRange.ColumnCount(); column++)
                    {
                        string cellValue = usedRange.Row(row).Cell(column).Value.ToString();
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            switch (usedRange.Row(1).Cell(column).Value.ToString())
                            {
                                case "Browser":
                                    browser = cellValue;
                                    break;
                                case "Page":
                                    string[] page = cellValue.Split(new[] { ":=" }, StringSplitOptions.RemoveEmptyEntries);
                                    //Creates XModule.
                                    if (cellValue.Contains(":="))
                                    {
                                        moduleId = Builder.CreateXModule(page[1], Definition.ModulesFolderId, new Dictionary<string, string> { { "Title", page[1] } });
                                    }
                                    else
                                    {
                                        browserName = browser;
                                        pageName = page[page.Length - 1];
                                        //If the Page properties are referred to Object Repository, it looks into the Object Repository XML for fetching the value.
                                        moduleId = Builder.CreateXModule(page[(page.Length-1)], Definition.ModulesFolderId, ObjectRepository.GetModuleProperties(objectrepositoryFolderPath, browserName, pageName));
                                    }
                                    //Creates XTestStep from XModule.
                                    testStepId = Builder.CreateXTestStepFromXModule(page[page.Length - 1], moduleId, testCaseId);
                                    break;
                                case "Object":
                                    //ObjectMap.GetBusinessType - This method returns the corresponding Tosca Business Type for UFT object classes.
                                    objectType = ObjectMap.GetBusinessType(cellValue);
                                    break;
                                case "Identifier":
                                    string[] objectIdentifier = cellValue.Split(new[] { ":=" }, StringSplitOptions.RemoveEmptyEntries);
                                    controlName = objectIdentifier[(objectIdentifier.Length - 1)];
                                    //ObjectMap.GetTechnicalId - This method returns the corresponding Tosca Technical ID Type for UFT object attributes.
                                    if (cellValue.Contains(":="))
                                    {
                                        //Control properties are defined in the excel file
                                        technicalIdParam.Add(ObjectMap.GetTechnicalId(objectIdentifier[0]), objectIdentifier[1]);
                                    }
                                    else
                                    {
                                        //Control properties referres to the Object Repository XML
                                        technicalIdParam = ObjectRepository.GetModuleAttributeProperties(objectrepositoryFolderPath, browserName, pageName,
                                                usedRange.Row(row).Cell(3).Value.ToString(), cellValue);
                                    }
                                    
                                    break;
                                case "Keyword":
                                    keyword = cellValue;
                                    break;
                                case "Value":
                                    data = cellValue;
                                    if (data.Contains("DT_") && !isTestCaseTemplate)
                                    {
                                        if (testSheetId == 0)
                                        {
                                            //The datasheet folder structure of the archive can be modified from here.
                                            testSheetId = ParseTestDataSheetFile(Directory.GetFiles(Definition.MigrationFolderPath + "Framework\\DataSheet")[0]);
                                        }
                                        //Converts TestCase to TestCaseTemplate
                                        Builder.ConvertTestCaseToTemplate(testCaseId, testSheetId);
                                        isTestCaseTemplate = true;
                                    }
                                    if (isTestCaseTemplate && data.Contains("DT_"))
                                    {
                                        string[] dataParts = data.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (dataParts.Length == 3)
                                        {
                                            data = dataParts[0] + ":" + dataParts[1] + ":{XL[" + dataParts[2].Replace("DT_", "") + "]}";
                                        }
                                        else if (dataParts.Length == 2)
                                        {
                                            data = dataParts[0] + ":{XL[" + dataParts[1].Replace("DT_", "") + "]}";
                                        }
                                        else
                                        {
                                            data = "{XL[" + data.Replace("DT_", "") + "]}";
                                        }
                                    }
                                    break;
                            }
                        }
                    }

                    int moduleAttributeId;
                    if (keyword == "Open")
                    {
                        //Creates SpecialExecutionTask (The 'TBox Open URL' module from Standard Subset is reused here).
                        moduleId = Builder.CreateSpecialExecutionTask("Open URL", Definition.ModulesFolderId, "Framework", "OpenUrl");
                        testStepId = Builder.CreateXTestStepFromXModule("Open URL", moduleId, testCaseId);
                        moduleAttributeId = Builder.CreateSpecialExecutionTaskAttribute("Url", moduleId);
                        Builder.SetXTestStepValue(data, testStepId, moduleAttributeId, null);
                    }
                    if (keyword == "CloseWindow")
                    {
                        //Creates SpecialExecutionTask (The 'TBox Window Operation' module from Standard Subset is reused here).
                        moduleId = Builder.CreateSpecialExecutionTask("TBox Window Operation", Definition.ModulesFolderId, "Framework", "WindowOperation");
                        testStepId = Builder.CreateXTestStepFromXModule("Close Window", moduleId, testCaseId);
                        moduleAttributeId = Builder.CreateSpecialExecutionTaskAttribute("Caption", moduleId);
                        Builder.SetXTestStepValue(data + "*", testStepId, moduleAttributeId, null);
                        moduleAttributeId = Builder.CreateSpecialExecutionTaskAttribute("Operation", moduleId);
                        Builder.SetXTestStepValue("Close", testStepId, moduleAttributeId, null);
                    }

                    if (technicalIdParam.Count == 0) continue;
                    if (objectType == "Table")
                    {
                        //Creates XModuleAttribute as a Table control (it adds the <Row>, <Col> and <Cell> attributes internally).
                        moduleAttributeId = Builder.CreateXModuleAttributeAsTable(controlName, moduleId, technicalIdParam, "", "");
                    }
                    else
                    {
                        //Creates XModuleAttribute (Page Controls)
                        moduleAttributeId = Builder.CreateXModuleAttribute(controlName, objectType, ActionMode.Input.ToString(), moduleId, technicalIdParam);
                    }

                    switch (keyword)
                    {
                        //Sets the XTestStepValue accroding to the Keyword.
                        case "Click":
                            Builder.SetXTestStepValue("{Click}", testStepId, moduleAttributeId, null);
                            break;
                        case "SetValue":
                            Builder.SetXTestStepValue(data, testStepId, moduleAttributeId, null);
                            break;
                        case "SelectValue":
                            Builder.SetXTestStepValue(data, testStepId, moduleAttributeId, null);
                            break;
                        case "GetValue":
                            Builder.SetXTestStepValue(data, testStepId, moduleAttributeId, ActionMode.Buffer.ToString());
                            break;
                        case "VerifyProperty":
                            string[] dataParts = data.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                            Builder.SetXTestStepValue(dataParts[0] + "=" + dataParts[1], testStepId, moduleAttributeId, ActionMode.Verify.ToString());
                            break;
                        case "VerifyProperty_Table":
                            dataParts = data.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                            Builder.SetXTestStepValueAsTableCell(dataParts[2], testStepId, moduleAttributeId, true, "#" + dataParts[0], "#" + dataParts[1], ActionMode.Verify.ToString());
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// This method migrates the excel sheet containing data.
        /// </summary>
        /// <param name="filePath">file path of the datasheet excel</param>
        /// <returns>The ID of the created TestSheet</returns>
        protected override int ParseTestDataSheetFile(string filePath)
        {
            //ClosedXML has been used here to read the excel files. The library can be found in Tosca Installation Directory at '%TRICENTIS_HOME%\ToscaCommander'.
            //Alternatively, Microsoft.Office.Interop.Excel library or other third-party library can also be used.
            XLWorkbook workbook = new XLWorkbook(filePath);
            IXLRange sheet = workbook.Worksheet(1).RangeUsed();
            int testSheetId = 0;
            int tcInstanceCollectionId = 0;
            Dictionary<string, int> attributeList = new Dictionary<string, int>();

            for (int row = 2; row <= sheet.RowCount(); row++)
            {
                int tcInstanceId = 0;
                for (int column = 1; column <= sheet.ColumnCount(); column++)
                {
                    string cellValue = sheet.Row(row).Cell(column).Value.ToString();
                    if (string.IsNullOrEmpty(cellValue)) continue;
                    switch (sheet.Row(1).Cell(column).Value.ToString())
                    {
                        case "TC_Name":
                            //Creates TestSheet
                            testSheetId = Builder.CreateTestSheet(cellValue, Definition.TestCaseDesignFolderId);
                            tcInstanceCollectionId = Builder.CreateInstanceCollection(testSheetId);
                            attributeList.Clear();
                            break;
                        case "Iteration":
                            //Creates Instance of the TestSheet
                            tcInstanceId = Builder.CreateInstance("TC_" + cellValue, tcInstanceCollectionId);
                            break;
                        default:
                            if (!attributeList.ContainsKey(sheet.Row(1).Cell(column).Value.ToString()))
                            {
                                //Creates the attribute of the TestSheet
                                int attributeId = Builder.CreateTDAttribute(sheet.Row(1).Cell(column).Value.ToString(), testSheetId);
                                attributeList.Add(sheet.Row(1).Cell(column).Value.ToString(), attributeId);
                            }
                            //Sets the attribute value for the corresponding Instance.
                            Builder.SetAttributeValue(cellValue, attributeList[sheet.Row(1).Cell(column).Value.ToString()], tcInstanceId);
                            break;
                    }
                }
            }
            return testSheetId;
        }
    }
}
