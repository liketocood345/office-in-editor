using System;
using System.Collections.Generic;
using System.Security;
using System.Text;

namespace OfficeInEditor.OfficeAddin
{
    internal static class RibbonMarkup
    {
        // IDs checked against Microsoft's office-fluent-ui-command-identifiers M365 lists.
        // A listed ID is an implemented target, not a claim of live Office-version acceptance testing.
        internal static readonly Dictionary<string, string[]> Contexts = new Dictionary<string, string[]>
        {
            ["word"] = new[] { "ContextMenuText", "ContextMenuHeading", "ContextMenuList",
                "ContextMenuTableCell", "ContextMenuTextTable" },
            ["excel"] = new[] { "ContextMenuCell", "ContextMenuListRange" },
            ["powerpoint"] = new[] { "ContextMenuTextEdit", "ContextMenuShape", "ContextMenuPicture",
                "ContextMenuTable", "ContextMenuThumbnail", "ContextMenuSlideSorter" }
        };
        internal static string Escape(string value) => SecurityElement.Escape(value ?? "");
        internal static string Build(string application)
        {
            var xml = new StringBuilder("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnRibbonLoad\"><ribbon><tabs><tab id=\"OieTab\" label=\"Office in Editor\"><group id=\"OieChatGroup\" label=\"Chat context\">");
            xml.Append("<button id=\"OieRefresh\" label=\"Refresh chats\" onAction=\"RefreshChats\" getEnabled=\"GetRefreshEnabled\"/>")
                .Append("<dynamicMenu id=\"OieTargets\" getLabel=\"GetTargetLabel\" getContent=\"GetChatMenu\"/>")
                .Append("<button id=\"OieSend\" label=\"Add to Chat\" onAction=\"AddToChat\" getEnabled=\"GetSendEnabled\"/>")
                .Append("<labelControl id=\"OieStatus\" getLabel=\"GetStatusLabel\"/>")
                .Append("</group></tab></tabs></ribbon><contextMenus>");
            if (Contexts.TryGetValue(application, out string[] contexts))
                for (int i = 0; i < contexts.Length; i++)
                    xml.Append("<contextMenu idMso=\"").Append(contexts[i]).Append("\"><button id=\"OieContext")
                        .Append(i).Append("\" label=\"Add to Chat\" insertBeforeMso=\"Cut\" onAction=\"AddToChat\" getEnabled=\"GetSendEnabled\"/></contextMenu>");
            return xml.Append("</contextMenus></customUI>").ToString();
        }
    }
}
