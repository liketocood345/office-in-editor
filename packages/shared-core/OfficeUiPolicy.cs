namespace OfficeInEditor.Core;

// Used only to exclude transient UI from legacy frame cleanup. Classification
// never authorizes changing a foreign window's parent, styles, focus or Z order.
internal static class OfficeUiPolicy
{
    internal static bool IsTransientUi(string name) =>
        name is "NetUIHWND" or "NUIDialog" or "MsoCommandBarPopup" or "MsoCommandBar"
        || name.StartsWith("MsoCommandBar", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("NetUI", StringComparison.OrdinalIgnoreCase);
}
