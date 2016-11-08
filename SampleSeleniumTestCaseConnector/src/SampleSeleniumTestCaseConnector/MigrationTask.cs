using System.Collections.Generic;
using ClosedXML.Excel;
using TCMigrationAPI;

namespace SampleSeleniumObjectMigratorAddOn
{
    /// <summary>
    /// This class contains the frmaework specific business logic for migration.
    /// </summary>
    public class MigrationTask
    {
        /// <summary>
        /// Definition object generates the migration metafile and required to access migration specific tasks from TCMigrationAPI.
        /// </summary>
        private ToscaObjectDefinition definition;
        /// <summary>
        /// Builder object provides utility methods to create Tosca Business Objects with required parameters.
        /// </summary>
        private ToscaObjectBuilder builder;

        /// <summary>
        /// Public Constructor
        /// </summary>
        /// <param name="toscaObjects">Definition Object</param>
        /// <param name="html">Engine Name</param>
        public MigrationTask(ToscaObjectDefinition toscaObjects, string html)
        {
            definition = toscaObjects;
            builder = new ToscaObjectBuilder(toscaObjects, html);
        }

        /// <summary>
        /// This method migrates the TestCase excel files.
        /// </summary>
        /// <param name="filepath">File path of the TestCase excel.</param>
        internal void ProcessTestScriptFile(string filepath)
        {
            //ClosedXML has been used here to read the excel files. The library can be found in Tosca Installation Directory at '%TRICENTIS_HOME%\ToscaCommander'.
            //Alternatively, Microsoft.Office.Interop.Excel library or other third-party library can also be used.
            XLWorkbook workbook = new XLWorkbook(filepath);

            foreach (IXLWorksheet sheet in workbook.Worksheets)
            {
                var xlRange = sheet.RangeUsed();
                int pagenumber = 1;
                //Builder object provides all the methods to create Tosca Business Objects. Here we create a TestCase. Also, Definition object provide the root folder IDs for import.
                int testCaseId = builder.CreateTestCase(sheet.Name, definition.TestCasesFolderId);
                int moduleId = 0;
                int testStepId = 0;
                bool newModuleRequired = true;
                for (int row = 2; row <= xlRange.RowCount(); row++)
                {
                    string keyword = null, locatorType = null, locatorValue = null, data = null;
                    for (int column = 1; column <= xlRange.ColumnCount(); column++)
                    {
                        string cellValue = xlRange.Row(row).Cell(column).Value.ToString();
                        switch (xlRange.Row(1).Cell(column).Value.ToString())
                        {
                            case "FunctionKeyword":
                                keyword = cellValue;
                                break;
                            case "Locator Type":
                                locatorType = cellValue;
                                break;
                            case "Locator Value":
                                locatorValue = cellValue;
                                break;
                            case "Test Data":
                                data = cellValue;
                                break;
                        }
                    }
                    int moduleAttributeId = 0;
                    if (keyword == "enter_URL")
                    {
                        //Creates SpecialExecutionTask (The 'Open URL' module from Standard Subset is reused here).
                        moduleId = builder.CreateSpecialExecutionTask("Open URL", definition.ModulesFolderId, "Framework", "OpenUrl");
                        testStepId = builder.CreateXTestStepFromXModule("Open URL", moduleId, testCaseId);
                        moduleAttributeId = builder.CreateSpecialExecutionTaskAttribute("Url", moduleId);
                        builder.SetXTestStepValue(data, testStepId, moduleAttributeId, null);
                        continue;
                    }
                    if (keyword == "close_Window")
                    {
                        //Creates SpecialExecutionTask (The 'Window Operation' module from Standard Subset is reused here).
                        moduleId = builder.CreateSpecialExecutionTask("TBox Window Operation", definition.ModulesFolderId, "Framework", "WindowOperation");
                        testStepId = builder.CreateXTestStepFromXModule("Close Window", moduleId, testCaseId);
                        moduleAttributeId = builder.CreateSpecialExecutionTaskAttribute("Caption", moduleId);
                        builder.SetXTestStepValue(data + "*", testStepId, moduleAttributeId, null);
                        moduleAttributeId = builder.CreateSpecialExecutionTaskAttribute("Operation", moduleId);
                        builder.SetXTestStepValue("Close", testStepId, moduleAttributeId, null);
                        continue;
                    }
                    switch (keyword)
                    {
                        case "click_On_Button":
                            if (moduleId == 0 || newModuleRequired)
                            {
                                moduleId = builder.CreateXModule("Page" + pagenumber, definition.ModulesFolderId, new Dictionary<string, string> { { "Title", "*" } });
                                testStepId = builder.CreateXTestStepFromXModule("Page" + pagenumber, moduleId, testCaseId);
                                newModuleRequired = false;
                            }
                            moduleAttributeId = builder.CreateXModuleAttribute(locatorValue, "Button",
                                ActionMode.Input.ToString(), moduleId,
                                new Dictionary<string, string> { { GetToscaType(locatorType), locatorValue } });
                            builder.SetXTestStepValue("{CLICK}", testStepId, moduleAttributeId, ActionMode.Input.ToString());

                            newModuleRequired = true;
                            ++pagenumber;
                            break;
                        case "select":
                            if (moduleId == 0 || newModuleRequired)
                            {
                                moduleId = builder.CreateXModule("Page" + pagenumber, definition.ModulesFolderId, new Dictionary<string, string> { { "Title", "*" } });
                                testStepId = builder.CreateXTestStepFromXModule("Page" + pagenumber, moduleId, testCaseId);
                                newModuleRequired = false;
                            }
                            moduleAttributeId = builder.CreateXModuleAttribute(locatorValue, "ComboBox",
                                ActionMode.Input.ToString(), moduleId,
                                new Dictionary<string, string> { { GetToscaType(locatorType), locatorValue } });
                            builder.SetXTestStepValue(data, testStepId, moduleAttributeId, ActionMode.Input.ToString());
                            break;
                        case "enter_Text":
                            if (moduleId == 0 || newModuleRequired)
                            {
                                moduleId = builder.CreateXModule("Page" + pagenumber, definition.ModulesFolderId, new Dictionary<string, string> { { "Title", "*" } });
                                testStepId = builder.CreateXTestStepFromXModule("Page" + pagenumber, moduleId, testCaseId);
                                newModuleRequired = false;
                            }
                            moduleAttributeId = builder.CreateXModuleAttribute(locatorValue, "TextBox",
                                ActionMode.Input.ToString(), moduleId,
                                new Dictionary<string, string> { { GetToscaType(locatorType), locatorValue } });
                            builder.SetXTestStepValue(data, testStepId, moduleAttributeId, ActionMode.Input.ToString());
                            break;

                    }
                }
            }
        }

        /// <summary>
        /// Gets Tosca Technical ID mapping
        /// </summary>
        /// <param name="locatorType">Selenium Locator Type</param>
        /// <returns>Tosca Technical ID name</returns>
        private string GetToscaType(string locatorType)
        {
            switch (locatorType)
            {
                case "id":
                    return "Id";
                default:
                    return null;
            }
        }
        
    }
}
