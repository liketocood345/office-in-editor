using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeInEditor.OfficeAddin
{
    internal sealed class DocumentRegistry
    {
        private sealed class Identity { internal readonly string Id = Guid.NewGuid().ToString("N"); }
        private readonly ConditionalWeakTable<object, Identity> identities = new ConditionalWeakTable<object, Identity>();
        private readonly Dictionary<string, object> bindings = new Dictionary<string, object>(StringComparer.Ordinal);
        private readonly dynamic application;
        private readonly string kind, officeId;
        internal DocumentRegistry(object application, string kind, string officeId)
        { this.application = application; this.kind = kind; this.officeId = officeId; }

        internal List<object> OpenDocuments()
        {
            dynamic collection;
            switch (kind)
            {
                case "word": collection = application.Documents; break;
                case "excel": collection = application.Workbooks; break;
                case "powerpoint": collection = application.Presentations; break;
                default: throw new BridgeException("UNSUPPORTED_APPLICATION");
            }
            int count = Convert.ToInt32(collection.Count);
            if (count > 128) throw new BridgeException("TOO_MANY_OPEN_DOCUMENTS");
            var result = new List<object>();
            for (int i = 1; i <= count; i++) result.Add(collection.Item(i));
            return result;
        }
        internal string Bind(object document)
        {
            string id = officeId + ":" + identities.GetValue(document, _ => new Identity()).Id;
            if (!bindings.ContainsKey(id) && bindings.Count >= 128)
            {
                var opened = OpenDocuments();
                var expired = new List<string>();
                foreach (var binding in bindings)
                    if (!opened.Exists(item => SameObject(item, binding.Value))) expired.Add(binding.Key);
                foreach (string key in expired) bindings.Remove(key);
                if (bindings.Count >= 128) throw new BridgeException("TOO_MANY_BOUND_DOCUMENTS");
            }
            bindings[id] = document;
            return id;
        }
        internal object Resolve(string id)
        {
            if (string.IsNullOrEmpty(id) || !bindings.TryGetValue(id, out object document))
                throw new BridgeException("UNKNOWN_DOCUMENT");
            // Never reopen by path/name or substitute ActiveDocument after a tab/focus change.
            if (!OpenDocuments().Exists(item => SameObject(item, document)))
            { bindings.Remove(id); throw new BridgeException("DOCUMENT_CLOSED"); }
            return document;
        }
        private static bool SameObject(object a, object b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (!Marshal.IsComObject(a) || !Marshal.IsComObject(b)) return false;
            IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
            try { first = Marshal.GetIUnknownForObject(a); second = Marshal.GetIUnknownForObject(b); return first == second; }
            finally { if (first != IntPtr.Zero) Marshal.Release(first); if (second != IntPtr.Zero) Marshal.Release(second); }
        }
        internal Dictionary<string, object> Describe(object documentObject)
        {
            dynamic document = documentObject;
            string name = Convert.ToString(document.Name, CultureInfo.InvariantCulture);
            if (name == null || name.Length > 256) throw new BridgeException("INVALID_DOCUMENT");
            string source = null;
            try
            {
                string fullName = Convert.ToString(document.FullName, CultureInfo.InvariantCulture);
                if (!string.IsNullOrEmpty(fullName) && fullName.Length <= 32767 && fullName.Length > 2 &&
                    char.IsLetter(fullName[0]) && fullName[1] == ':' && (fullName[2] == '\\' || fullName[2] == '/'))
                    source = Path.GetFullPath(fullName);
            }
            catch { }
            bool dirty = true, enabled = false;
            try { dirty = !Convert.ToBoolean(document.Saved, CultureInfo.InvariantCulture); } catch { }
            string state;
            try { enabled = Convert.ToBoolean(document.AutoSaveOn, CultureInfo.InvariantCulture); state = enabled ? "enabled" : "off_or_unavailable"; }
            catch { state = "unsupported"; }
            return Json.Object("application", kind, "documentId", Bind(documentObject), "documentName", name,
                "sourcePath", source, "dirty", dirty, "autoSaveAvailable", enabled ? (object)true : null,
                "autoSaveEnabled", enabled, "autoSaveState", state, "nativeAutoSaveLocalRedirect", false);
        }
        internal void Clear() => bindings.Clear();
    }
}
