using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using Opc.Ua;

namespace OpcUaExcelAddin
{
    public sealed class TagCache
    {
        private static readonly Lazy<TagCache> LazyInstance = new Lazy<TagCache>(() => new TagCache());
        private readonly ConcurrentDictionary<string, object> _values = new ConcurrentDictionary<string, object>(StringComparer.Ordinal);

        public static TagCache Instance => LazyInstance.Value;

        public event Action<string, object> ValueChanged;

        public object GetValue(string nodeId)
        {
            return _values.TryGetValue(nodeId, out object value) ? value : "#NOT_CONNECTED";
        }

        public void SetValue(string nodeId, object value)
        {
            if (string.IsNullOrEmpty(nodeId))
            {
                return;
            }

            object normalizedValue = NormalizeValue(value);
            while (true)
            {
                if (_values.TryGetValue(nodeId, out object previousValue))
                {
                    if (Equals(previousValue, normalizedValue))
                    {
                        return;
                    }

                    if (_values.TryUpdate(nodeId, normalizedValue, previousValue))
                    {
                        NotifyValueChanged(nodeId, normalizedValue);
                        return;
                    }
                }
                else if (_values.TryAdd(nodeId, normalizedValue))
                {
                    NotifyValueChanged(nodeId, normalizedValue);
                    return;
                }
            }
        }

        public void SetDataValue(string nodeId, DataValue value)
        {
            if (value == null)
            {
                SetValue(nodeId, "#ERROR");
                return;
            }

            if (StatusCode.IsBad(value.StatusCode))
            {
                if (value.StatusCode == StatusCodes.BadNodeIdInvalid)
                {
                    SetValue(nodeId, "#BAD_NODEID");
                }
                else if (value.StatusCode == StatusCodes.BadNodeIdUnknown || value.StatusCode == StatusCodes.BadMonitoredItemIdInvalid)
                {
                    SetValue(nodeId, "#NODE_NOT_FOUND");
                }
                else
                {
                    SetValue(nodeId, "#ERROR");
                }

                return;
            }

            SetValue(nodeId, value.Value);
        }

        private static object NormalizeValue(object value)
        {
            if (value == null)
            {
                return "#ERROR";
            }

            if (value is string || value is bool || value is DateTime)
            {
                return value;
            }

            if (value is byte[] bytes)
            {
                return Convert.ToBase64String(bytes);
            }

            TypeCode typeCode = Type.GetTypeCode(value.GetType());
            switch (typeCode)
            {
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
            }

            return value.ToString();
        }

        private void NotifyValueChanged(string nodeId, object value)
        {
            Action<string, object> handlers = ValueChanged;
            if (handlers == null)
            {
                return;
            }

            foreach (Action<string, object> handler in handlers.GetInvocationList())
            {
                try
                {
                    handler(nodeId, value);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Tag cache subscriber failed: " + ex.Message);
                }
            }
        }
    }
}