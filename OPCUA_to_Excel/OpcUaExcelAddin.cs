using System;
using System.Diagnostics;
using ExcelDna.Integration;

namespace OpcUaExcelAddin
{
    public sealed class OpcUaExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ConfigLoader.Load();
            ApplyRtdThrottleInterval();
            SubscriptionManager.Instance.Initialize(OpcUaService.Instance);
            OpcUaService.Instance.Start();
        }

        public void AutoClose()
        {
            SubscriptionManager.Instance.Stop();
            OpcUaService.Instance.Disconnect();
        }

        private static void ApplyRtdThrottleInterval()
        {
            try
            {
                dynamic excelApplication = ExcelDnaUtil.Application;
                excelApplication.RTD.ThrottleInterval = ConfigLoader.RtdThrottleIntervalMs;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Could not set Excel RTD throttle interval: " + ex.Message);
            }
        }
    }
}