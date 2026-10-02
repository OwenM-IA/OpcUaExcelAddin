# OPC UA Excel Add-in

## Install in Excel

1. Copy the entire `publish` folder to a location you can keep, such as Documents. Keep `appsettings.json` and `tags.json` beside the `.xll` files. You do not need to copy the files into Excel's installation folder.
2. In Excel, check your version's bitness at **File > Account > About Excel**.
3. Go to **File > Options > Add-ins**. At the bottom, set **Manage** to **Excel Add-ins**, click **Go**, then **Browse**.
4. Select the matching add-in:
   - 64-bit Excel: `OpcUaExcelAddin.xll`
   - 32-bit Excel: `OpcUaExcelAddin-x86.xll`
5. Select the add-in in the list and click **OK**.

If you move the folder later, browse to the `.xll` at its new location. Keep the JSON files beside it.

## Read a tag

Enter a formula in a worksheet cell, replacing the example with a NodeId available on your OPC UA server:

```excel
=OPCUA("ns=6;s=_gOBit")
```

The computer running Excel must be able to reach the OPC UA server address configured in `appsettings.json`. The add-in tries configured addresses in order and connects to the first available one.

In Excel, the add-in's RTD server is identified as `OpcUaExcelAddin.LiveTags`.

In `appsettings.json`, `SamplingIntervalMs` controls how often the OPC UA server is asked to sample a tag, and `PublishingIntervalMs` controls how often it is asked to publish subscription updates. `RtdThrottleIntervalMs` controls how often Excel refreshes displayed RTD values; it defaults to 1000 milliseconds. The sampling and publishing defaults are 500 milliseconds. The OPC UA server may revise its requested intervals. Lower values can increase server, network, and Excel load. After changing these settings, close Excel completely and reopen it.

## Optional aliases

`tags.json` maps friendly names to NodeIds. For example:

```json
{
  "MachineSpeed": "ns=6;s=MachineSpeed"
}
```

Then enter `=OPCUA("MachineSpeed")` in Excel instead of the full NodeId.

After changing `appsettings.json` or `tags.json`, close Excel completely and reopen it so the add-in reloads the settings.

## Troubleshooting

- `#NOT_CONNECTED`: Excel cannot connect to a configured OPC UA server. Check network access and the endpoint in `appsettings.json`.
- `#NODE_NOT_FOUND`: Check that the NodeId and namespace index are correct for that server.

The add-in currently accepts untrusted OPC UA server certificates automatically. Only use it with servers and networks you trust.
