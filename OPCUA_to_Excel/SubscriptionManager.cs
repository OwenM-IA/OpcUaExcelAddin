using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Opc.Ua;
using Opc.Ua.Client;

namespace OpcUaExcelAddin
{
    public sealed class SubscriptionManager
    {
        private static readonly Lazy<SubscriptionManager> LazyInstance = new Lazy<SubscriptionManager>(() => new SubscriptionManager());
        private readonly SemaphoreSlim _subscriptionGate = new SemaphoreSlim(1, 1);
        private readonly object _requestedLock = new object();
        private readonly HashSet<string> _requestedNodeIds = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _activeNodeIds = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<string, MonitoredItem> _monitoredItems = new Dictionary<string, MonitoredItem>(StringComparer.Ordinal);
        private OpcUaService _service;
        private Session _session;
        private Subscription _subscription;

        private SubscriptionManager()
        {
        }

        public static SubscriptionManager Instance => LazyInstance.Value;

        public void Initialize(OpcUaService service)
        {
            if (ReferenceEquals(_service, service))
            {
                return;
            }

            if (_service != null)
            {
                _service.Connected -= OnConnected;
                _service.Disconnected -= OnDisconnected;
            }

            _service = service;
            _service.Connected += OnConnected;
            _service.Disconnected += OnDisconnected;
        }

        public async Task EnsureSubscribedAsync(string nodeId)
        {
            if (!ConfigLoader.TryResolveNodeId(nodeId, out string resolvedNodeId))
            {
                TagCache.Instance.SetValue(nodeId, "#BAD_NODEID");
                return;
            }

            lock (_requestedLock)
            {
                _requestedNodeIds.Add(resolvedNodeId);
            }

            await _subscriptionGate.WaitAsync().ConfigureAwait(false);
            try
            {
                await SubscribeCoreAsync(resolvedNodeId).ConfigureAwait(false);
            }
            finally
            {
                _subscriptionGate.Release();
            }
        }

        public void Stop()
        {
            if (_service != null)
            {
                _service.Connected -= OnConnected;
                _service.Disconnected -= OnDisconnected;
                _service = null;
            }

            _subscriptionGate.Wait();
            try
            {
                _subscription?.Dispose();
                _subscription = null;
                _session = null;
                _activeNodeIds.Clear();
                _monitoredItems.Clear();
                lock (_requestedLock)
                {
                    _requestedNodeIds.Clear();
                }
            }
            finally
            {
                _subscriptionGate.Release();
            }
        }

        public async Task ReleaseTagAsync(string nodeId)
        {
            lock (_requestedLock)
            {
                _requestedNodeIds.Remove(nodeId);
            }

            await _subscriptionGate.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_monitoredItems.TryGetValue(nodeId, out MonitoredItem monitoredItem))
                {
                    monitoredItem.Notification -= OnNotification;
                    _subscription.RemoveItem(monitoredItem);
                    _subscription.ApplyChanges();
                    _monitoredItems.Remove(nodeId);
                    _activeNodeIds.Remove(nodeId);
                }

                TagCache.Instance.SetValue(nodeId, "#NOT_CONNECTED");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OPC UA monitored item removal failed: " + ex);
            }
            finally
            {
                _subscriptionGate.Release();
            }
        }

        private async Task SubscribeCoreAsync(string nodeId)
        {
            Session session = _service == null ? null : _service.CurrentSession;
            if (session == null || !_service.IsConnected)
            {
                TagCache.Instance.SetValue(nodeId, "#NOT_CONNECTED");
                return;
            }

            if (!ReferenceEquals(_session, session))
            {
                _subscription?.Dispose();
                _subscription = null;
                _session = session;
                _activeNodeIds.Clear();
                _monitoredItems.Clear();
            }

            if (_activeNodeIds.Contains(nodeId))
            {
                return;
            }

            try
            {
                NodeId parsedNodeId = NodeId.Parse(nodeId);
                DataValue initialValue = session.ReadValue(parsedNodeId);
                TagCache.Instance.SetDataValue(nodeId, initialValue);
                if (StatusCode.IsBad(initialValue.StatusCode))
                {
                    return;
                }

                if (_subscription == null)
                {
                    _subscription = new Subscription(session.DefaultSubscription)
                    {
                        DisplayName = "OpcUaExcelAddin",
                        PublishingInterval = ConfigLoader.PublishingIntervalMs,
                        KeepAliveCount = 10,
                        LifetimeCount = 100,
                        MaxNotificationsPerPublish = 1000,
                        PublishingEnabled = true,
                        Priority = 100
                    };
                    session.AddSubscription(_subscription);
                    _subscription.Create();
                }

                var monitoredItem = new MonitoredItem(_subscription.DefaultItem)
                {
                    StartNodeId = parsedNodeId,
                    AttributeId = Attributes.Value,
                    DisplayName = nodeId,
                    SamplingInterval = ConfigLoader.SamplingIntervalMs,
                    QueueSize = 1,
                    DiscardOldest = true
                };
                monitoredItem.Notification += OnNotification;
                _subscription.AddItem(monitoredItem);
                _subscription.ApplyChanges();
                _activeNodeIds.Add(nodeId);
                _monitoredItems.Add(nodeId, monitoredItem);
            }
            catch (ServiceResultException ex)
            {
                if (ex.StatusCode == StatusCodes.BadNodeIdInvalid)
                {
                    TagCache.Instance.SetValue(nodeId, "#BAD_NODEID");
                }
                else if (ex.StatusCode == StatusCodes.BadNodeIdUnknown)
                {
                    TagCache.Instance.SetValue(nodeId, "#NODE_NOT_FOUND");
                }
                else
                {
                    Debug.WriteLine("OPC UA subscription failed: " + ex);
                    TagCache.Instance.SetValue(nodeId, _service != null && _service.IsConnected ? "#ERROR" : "#NOT_CONNECTED");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OPC UA subscription failed: " + ex);
                TagCache.Instance.SetValue(nodeId, _service != null && _service.IsConnected ? "#ERROR" : "#NOT_CONNECTED");
            }
        }

        private void OnConnected(Session session)
        {
            Task.Run(async () =>
            {
                await _subscriptionGate.WaitAsync().ConfigureAwait(false);
                try
                {
                    _subscription?.Dispose();
                    _subscription = null;
                    _session = session;
                    _activeNodeIds.Clear();
                    _monitoredItems.Clear();

                    string[] nodeIds;
                    lock (_requestedLock)
                    {
                        nodeIds = new string[_requestedNodeIds.Count];
                        _requestedNodeIds.CopyTo(nodeIds);
                    }

                    foreach (string nodeId in nodeIds)
                    {
                        await SubscribeCoreAsync(nodeId).ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("OPC UA resubscription failed: " + ex);
                }
                finally
                {
                    _subscriptionGate.Release();
                }
            });
        }

        private void OnDisconnected()
        {
            string[] nodeIds;
            lock (_requestedLock)
            {
                nodeIds = new string[_requestedNodeIds.Count];
                _requestedNodeIds.CopyTo(nodeIds);
            }

            foreach (string nodeId in nodeIds)
            {
                TagCache.Instance.SetValue(nodeId, "#NOT_CONNECTED");
            }
        }

        private static void OnNotification(MonitoredItem monitoredItem, MonitoredItemNotificationEventArgs args)
        {
            foreach (DataValue value in monitoredItem.DequeueValues())
            {
                TagCache.Instance.SetDataValue(monitoredItem.DisplayName, value);
            }
        }
    }
}