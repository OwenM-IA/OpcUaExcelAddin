using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Newtonsoft.Json;
using Opc.Ua;

namespace OpcUaExcelAddin
{
    public static class ConfigLoader
    {
        private const int DefaultSamplingIntervalMs = 500;
        private const int DefaultPublishingIntervalMs = 500;
        private const int DefaultRtdThrottleIntervalMs = 1000;
        private static readonly object SyncRoot = new object();
        private static Dictionary<string, string> _aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static string[] _endpointUrls = new string[0];
        private static int _samplingIntervalMs = DefaultSamplingIntervalMs;
        private static int _publishingIntervalMs = DefaultPublishingIntervalMs;
        private static int _rtdThrottleIntervalMs = DefaultRtdThrottleIntervalMs;
        private static volatile bool _loaded;

        public static string EndpointUrl
        {
            get
            {
                EnsureLoaded();
                return _endpointUrls.Length == 0 ? null : _endpointUrls[0];
            }
        }

        public static string[] EndpointUrls
        {
            get
            {
                EnsureLoaded();
                return (string[])_endpointUrls.Clone();
            }
        }

        public static int SamplingIntervalMs
        {
            get
            {
                EnsureLoaded();
                return _samplingIntervalMs;
            }
        }

        public static int PublishingIntervalMs
        {
            get
            {
                EnsureLoaded();
                return _publishingIntervalMs;
            }
        }

        public static int RtdThrottleIntervalMs
        {
            get
            {
                EnsureLoaded();
                return _rtdThrottleIntervalMs;
            }
        }

        public static string ConfigurationDirectory
        {
            get
            {
                string xllPath = ExcelDnaUtil.XllPath;
                if (!string.IsNullOrWhiteSpace(xllPath))
                {
                    return Path.GetDirectoryName(xllPath);
                }

                return AppDomain.CurrentDomain.BaseDirectory;
            }
        }

        public static void Load()
        {
            lock (SyncRoot)
            {
                if (_loaded)
                {
                    return;
                }

                string directory = ConfigurationDirectory;
                string settingsPath = Path.Combine(directory, "appsettings.json");
                string tagsPath = Path.Combine(directory, "tags.json");

                try
                {
                    if (File.Exists(settingsPath))
                    {
                        var settings = JsonConvert.DeserializeObject<AddinSettings>(File.ReadAllText(settingsPath));
                        if (settings != null)
                        {
                            _samplingIntervalMs = settings.SamplingIntervalMs > 0
                                ? settings.SamplingIntervalMs
                                : DefaultSamplingIntervalMs;
                            _publishingIntervalMs = settings.PublishingIntervalMs > 0
                                ? settings.PublishingIntervalMs
                                : DefaultPublishingIntervalMs;
                            _rtdThrottleIntervalMs = settings.RtdThrottleIntervalMs > 0
                                ? settings.RtdThrottleIntervalMs
                                : DefaultRtdThrottleIntervalMs;

                            string[] configuredEndpoints = settings.EndpointUrls;
                            if (configuredEndpoints == null || configuredEndpoints.Length == 0)
                            {
                                configuredEndpoints = new[] { settings.EndpointUrl };
                            }

                            var endpointUrls = new List<string>();
                            foreach (string endpointUrl in configuredEndpoints)
                            {
                                if (!string.IsNullOrWhiteSpace(endpointUrl) &&
                                    !endpointUrls.Contains(endpointUrl.Trim(), StringComparer.OrdinalIgnoreCase))
                                {
                                    endpointUrls.Add(endpointUrl.Trim());
                                }
                            }

                            _endpointUrls = endpointUrls.ToArray();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Could not load appsettings.json: " + ex.Message);
                }

                try
                {
                    if (File.Exists(tagsPath))
                    {
                        var aliases = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(tagsPath));
                        if (aliases != null)
                        {
                            _aliases = new Dictionary<string, string>(aliases, StringComparer.OrdinalIgnoreCase);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Could not load tags.json: " + ex.Message);
                }

                _loaded = true;
            }
        }

        public static bool TryResolveNodeId(string tagOrAlias, out string nodeId)
        {
            EnsureLoaded();
            nodeId = null;

            if (string.IsNullOrWhiteSpace(tagOrAlias))
            {
                return false;
            }

            string input = tagOrAlias.Trim();
            if (_aliases.TryGetValue(input, out string aliasNodeId))
            {
                input = aliasNodeId;
            }

            try
            {
                NodeId.Parse(input);
                nodeId = input;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void EnsureLoaded()
        {
            if (!_loaded)
            {
                Load();
            }
        }

        private sealed class AddinSettings
        {
            public string EndpointUrl { get; set; }
            public string[] EndpointUrls { get; set; }
            public int SamplingIntervalMs { get; set; }
            public int PublishingIntervalMs { get; set; }
            public int RtdThrottleIntervalMs { get; set; }
        }
    }
}