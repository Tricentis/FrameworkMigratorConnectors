using System;
using System.IO;
using TCMigrationAPI;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;

namespace SampleUFTTestCaseConnector
{
    public class UFTMigrationAddOnTask : TCAddOnTask
    {
        /// <summary>
        /// This method contains the logic performed when the context menu is selected.
        /// </summary>
        /// <param name="objectToExecuteOn">TCObject on which the context menu task is performed.</param>
        /// <param name="taskContext">Task Context of the AddOn Task.</param>
        /// <returns>TCObject Instance</returns>
        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            TCProject workspaceRoot = objectToExecuteOn as TCProject;
            //Opens the File Upload dialog.
            string filepath = taskContext.GetFilePath("Upload UFT Object archive");
            //Instantiation of this object is mandatory. This class contains the necessary methods for migration.
            ToscaObjectDefinition toscaObjects = new ToscaObjectDefinition();
            try
            {
                //Instantiates the MigrationTask class that contains the business logic of migration.
                MigrationTask migrationObjectImporter = new MigrationTask(toscaObjects, Engine.Html);
                //Entry point of MigrationTask class. 
                migrationObjectImporter.ProcessArchive(filepath);

                //Calling this method is mandatory. It outputs the file containing the migrated object information.
                string outputFilePath = toscaObjects.FinishObjectDefinitionTask();
                //Imports the output file from MigrationTask.
                workspaceRoot?.ImportExternalObjects(outputFilePath);
                //Cleans the migration metafiles.
                Directory.Delete(toscaObjects.MigrationFolderPath, true);
            }
            catch (Exception e)
            {
                //Pops-up the error message in case of any error in Migration.
                taskContext.ShowErrorMessage("Exception occured", e.Message);
            }
            return null;
        }

        /// <summary>
        /// This sets the name of the context menu sub-option.
        /// </summary>
        public override string Name { get { return "Import UFT Test Case"; } }

        /// <summary>
        /// This sets the type of Business Object the context menu will be available on (in this case, the workspace root).
        /// </summary>
        public override Type ApplicableType { get { return typeof (TCProject); } }
    }
}
