using Tricentis.TCAddOns;

namespace SampleManualTestCaseConnector
{
    public class ManualTCInportAddOn : TCAddOn
    {
        /// <summary>
        /// This sets the name of the context menu option.
        /// </summary>
        public override string UniqueName { get { return "Framework Migration"; } }
    }
}
