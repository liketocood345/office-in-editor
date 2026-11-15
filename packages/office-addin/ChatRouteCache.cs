using System;
using System.Collections.Generic;

namespace OfficeInEditor.OfficeAddin
{
    // Display/cache only. Core resolves the authoritative destination at capture.
    internal sealed class ChatRouteCache
    {
        internal sealed class Route
        {
            internal string Id, Label, Epoch, SessionId;
            internal bool Manual;
            internal long Version;
        }
        private readonly Dictionary<string, Route> documents = new Dictionary<string, Route>(StringComparer.Ordinal);
        private readonly HashSet<string> retiredBrokers = new HashSet<string>(StringComparer.Ordinal);
        private string broker;
        internal Route For(string documentId)
        {
            if (string.IsNullOrEmpty(documentId)) return new Route();
            if (!documents.TryGetValue(documentId, out Route route)) documents.Add(documentId, route = new Route());
            return route;
        }
        internal bool Bind(string documentId, string sessionId, string id, string label, string epoch, long version, string brokerEpoch)
        {
            if (string.IsNullOrEmpty(documentId) || string.IsNullOrEmpty(id) || string.IsNullOrEmpty(epoch) ||
                string.IsNullOrEmpty(brokerEpoch) || version < 1 || retiredBrokers.Contains(brokerEpoch)) return false;
            if (broker != brokerEpoch)
            {
                if (broker != null) retiredBrokers.Add(broker);
                broker = brokerEpoch; documents.Clear();
            }
            var route = For(documentId);
            if (version <= route.Version) return false;
            route.Id = id; route.Label = label; route.Epoch = epoch; route.SessionId = sessionId;
            route.Version = version; route.Manual = false;
            return true;
        }
        internal void Clear() { documents.Clear(); retiredBrokers.Clear(); broker = null; }
    }
}
