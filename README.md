# OpcUaExcelAddin

Excel-DNA add-in for reading live OPC UA values in Excel through subscriptions and RTD.

## Requirements

- Windows with Visual Studio 2022 and the .NET Framework 4.8 Developer Pack
- Microsoft Excel; use the add-in matching Excel's bitness
- Network access to an OPC UA server

## Configure

Edit `appsettings.json` and set `EndpointUrls` to an ordered list of OPC UA server endpoints. The add-in tries each endpoint in order and connects to the first available one; the same order is used for automatic reconnect. The legacy single `EndpointUrl` setting is also supported. `tags.json` is optional; keys are formula aliases and values are valid OPC UA NodeIds. The sample files are copied beside the packed add-ins during build.

The requested `SamplingIntervalMs` and `PublishingIntervalMs` are configurable in `appsettings.json`; both default to 500 milliseconds when omitted or set to a non-positive value. The OPC UA server may revise these requested intervals.

Examples:

```excel
=OPCUA("ns=6;s=MyTag")
=OPCUA("MachineSpeed")
```

The first formula use creates a monitored item. Subsequent recalculations return the cached value; OPC UA data-change notifications update the cache and Excel RTD topic. Supported OPC UA scalar values are returned as Excel values; other values are returned as text.

The add-in connects automatically when loaded and retries the connection every five seconds after a failure. Error results are returned as `#NOT_CONNECTED`, `#NODE_NOT_FOUND`, `#BAD_NODEID`, or `#ERROR` text.

## Build

1. Open `OpcUaExcelAddin.sln` in Visual Studio 2022.
2. Restore NuGet packages and select **Release**.
3. Build the solution.

The Excel-DNA build targets generate packed files under `bin\Release\publish\`. `OpcUaExcelAddin.xll` is the 64-bit add-in; `OpcUaExcelAddin-x86.xll` is for 32-bit Excel. A matching Excel-DNA `.dna` file is also generated in the build output.

## Publish and install

Copy the matching `.xll`, `appsettings.json`, and `tags.json` from `bin\Release\publish\` to a deployment folder. In Excel, open **File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse**, select the `.xll`, and enable it. Update the endpoint and aliases in the JSON files before distributing the folder.

## Security note

Untrusted OPC UA server certificates are automatically accepted to meet the requested behavior. The client also creates and stores its application certificate under `%LOCALAPPDATA%\OpcUaExcelAddin\pki`. Automatic certificate acceptance removes server identity verification; use only on a trusted network, and switch to certificate validation before production deployment.
