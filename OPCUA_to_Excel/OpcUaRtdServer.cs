using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ExcelDna.Integration.Rtd;

namespace OpcUaExcelAddin
{
    [ComVisible(true)]
    public sealed class OpcUaRtdServer : ExcelRtdServer
    {
        private readonly ConcurrentDictionary<string, ConcurrentDictionary<int, Topic>> _topics =
            new ConcurrentDictionary<string, ConcurrentDictionary<int, Topic>>();

        protected override bool ServerStart()
        {
            TagCache.Instance.ValueChanged += OnValueChanged;
            return true;
        }

        protected override object ConnectData(Topic topic, System.Collections.Generic.IList<string> topicInfo, ref bool newValues)
        {
            if (topicInfo == null || topicInfo.Count == 0 || !ConfigLoader.TryResolveNodeId(topicInfo[0], out string nodeId))
            {
                newValues = true;
                return "#BAD_NODEID";
            }

            ConcurrentDictionary<int, Topic> nodeTopics = _topics.GetOrAdd(
                nodeId,
                key => new ConcurrentDictionary<int, Topic>());
            nodeTopics[topic.TopicId] = topic;
            newValues = true;

            object currentValue = TagCache.Instance.GetValue(nodeId);
            _ = Task.Run(() => SubscriptionManager.Instance.EnsureSubscribedAsync(nodeId));
            return currentValue;
        }

        protected override void DisconnectData(Topic topic)
        {
            foreach (var pair in _topics)
            {
                if (pair.Value.TryRemove(topic.TopicId, out Topic ignored) && pair.Value.IsEmpty)
                {
                    _topics.TryRemove(pair.Key, out ConcurrentDictionary<int, Topic> removed);
                    _ = ReleaseWhenUnusedAsync(pair.Key);
                }
            }
        }

        protected override void ServerTerminate()
        {
            TagCache.Instance.ValueChanged -= OnValueChanged;
        }

        private void OnValueChanged(string nodeId, object value)
        {
            if (_topics.TryGetValue(nodeId, out ConcurrentDictionary<int, Topic> nodeTopics))
            {
                foreach (Topic topic in nodeTopics.Values)
                {
                    topic.UpdateValue(value);
                }
            }
        }

        private async Task ReleaseWhenUnusedAsync(string nodeId)
        {
            await Task.Delay(250).ConfigureAwait(false);
            if (!_topics.ContainsKey(nodeId))
            {
                await SubscriptionManager.Instance.ReleaseTagAsync(nodeId).ConfigureAwait(false);
            }
        }
    }
}