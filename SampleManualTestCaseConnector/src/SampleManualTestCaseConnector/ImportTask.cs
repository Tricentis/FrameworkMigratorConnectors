using ClosedXML.Excel;
using TCMigrationAPI;

namespace SampleManualTestCaseConnector
{
    /// <summary>
    /// This class contains the business logic for the igration of Manual Test Cases.
    /// </summary>
    public class ImportTask
    {
        /// <summary>
        /// Definition object generates the migration metafile and required to access migration specific tasks from TCMigrationAPI.
        /// </summary>
        private ToscaObjectDefinition Definition;
        /// <summary>
        /// Builder object provides utility methods to create Tosca Business Objects with required parameters.
        /// </summary>
        private ToscaObjectBuilder Builder;

        /// <summary>
        /// Public Constructor
        /// </summary>
        /// <param name="definition">Definition Object</param>
        public ImportTask(ToscaObjectDefinition definition)
        {
            Definition = definition;
            //The 'Engine' parameter in ToscaObjectBuilder constructor is passed as null because we doen't need any specific engine for ManualTestCase.
            Builder = new ToscaObjectBuilder(Definition, null);
        }

        /// <summary>
        /// This method handles the migration of Manual TestCase
        /// </summary>
        /// <param name="filePath">File Path of the ManualTestCase Excel sheet.</param>
        public void ProcessManualTestCaseFile(string filePath)
        {
            //ClosedXML has been used here to read the excel files. The library can be found in Tosca Installation Directory at '%TRICENTIS_HOME%\ToscaCommander'.
            //Alternatively, Microsoft.Office.Interop.Excel library or other third-party library can also be used.
            XLWorkbook workBook = new XLWorkbook(filePath);

            foreach (IXLWorksheet sheet in workBook.Worksheets)
            {
                IXLRange usedRange = sheet.RangeUsed();
                int testCaseId = 0;
                int testStepId = 0;
                for (int row = 2; row <= usedRange.RowCount(); row++)
                {
                    for (int column = 1; column <= usedRange.ColumnCount(); column++)
                    {
                        string cellValue = usedRange.Row(row).Cell(column).Value.ToString();
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            switch (usedRange.Row(1).Cell(column).Value.ToString())
                            {
                                case "TestCase":
                                    //Creates TestCase
                                    testCaseId = Builder.CreateTestCase(cellValue, Definition.TestCasesFolderId);
                                    break;
                                case "Action":
                                    //Creates ManualTestStep
                                    testStepId = Builder.CreateManualTestStep(cellValue, testCaseId, null);
                                    break;
                                case "Input Parameter":
                                    string value = string.IsNullOrEmpty(usedRange.Row(row).Cell(column+1).Value.ToString())
                                        ? ""
                                        : usedRange.Row(row).Cell(column+1).Value.ToString();
                                    //Creates ManualTestStepValue with ActionMode as Input
                                    Builder.CreateManualTestStepValue(cellValue, testStepId, value, ActionMode.Input.ToString(), null);
                                    break;
                                case "Expected Result":
                                    //Creates ManualTestStepValue with ActionMode as Verify
                                    Builder.CreateManualTestStepValue(cellValue, testStepId, "", ActionMode.Verify.ToString(), null);
                                    break;
                            }
                        }
                    }
                }
            }
        }
    }
}
