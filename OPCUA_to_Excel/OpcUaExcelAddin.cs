using ExcelDna.Integration;

namespace OpcUaExcelAddin
{
    public sealed class OpcUaExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ConfigLoader.Load();
            SubscriptionManager.Instance.Initialize(OpcUaService.Instance);
            OpcUaService.Instance.Start();
        }

        public void AutoClose()
        {
            SubscriptionManager.Instance.Stop();
            OpcUaService.Instance.Disconnect();
        }
    }
}