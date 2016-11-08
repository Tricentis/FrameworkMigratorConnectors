using Tricentis.TCAddOns;

namespace SampleUFTTestCaseConnector
{
    public class UFTMigrationAddOn : TCAddOn
    {
        /// <summary>
        /// This sets the name of the context menu option.
        /// </summary>
        public override string UniqueName { get { return "Framework Migration"; } }
    }
}
