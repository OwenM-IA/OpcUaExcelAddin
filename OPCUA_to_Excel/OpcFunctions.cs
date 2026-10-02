using System;
using ExcelDna.Integration;
using Opc.Ua;

namespace OpcUaExcelAddin
{
    public static class OpcFunctions
    {
        [ExcelFunction(Name = "OPCUA", Description = "Returns the latest subscribed OPC UA node value.")]
        public static object OPCUA(
            [ExcelArgument(Description = "An OPC UA NodeId or an alias from tags.json.")] string tag)
        {
            if (!ConfigLoader.TryResolveNodeId(tag, out string nodeId))
            {
                return "#BAD_NODEID";
            }

            try
            {
                return XlCall.RTD(OpcUaRtdServer.ServerProgId, null, nodeId);
            }
            catch
            {
                return "#ERROR";
            }
        }
    }
}