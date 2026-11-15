using System;
using System.Collections.Generic;
using System.Security;
using System.Text;

namespace OfficeInEditor.OfficeAddin
{
    internal static class RibbonMarkup
    {
        // Word LTSC 2021: only ContextMenuText is attached. Extra heading/table IDs from M365
        // lists can make Word reject the whole customUI and disable the add-in.
        internal static readonly Dictionary<string, string[]> Contexts = new Dictionary<string, string[]>
        {
            ["word"] = new[] { "ContextMenuText" },
            ["excel"] = new[] { "ContextMenuCell", "ContextMenuListRange" },
            ["powerpoint"] = new[] { "ContextMenuTextEdit", "ContextMenuShape", "ContextMenuPicture",
                "ContextMenuTable", "ContextMenuThumbnail", "ContextMenuSlideSorter" }
        };
        internal static string Escape(string value) => SecurityElement.Escape(value ?? "");

        internal static string FromRibbonId(string ribbonId)
        {
            if (string.IsNullOrEmpty(ribbonId)) return "unsupported";
            if (ribbonId.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0) return "word";
            if (ribbonId.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0) return "excel";
            if (ribbonId.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0) return "powerpoint";
            return "unsupported";
        }

        internal static string Build(string application)
        {
            var xml = new StringBuilder("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnRibbonLoad\">");
            xml.Append("<commands><command idMso=\"FileSave\" getEnabled=\"GetFileSaveEnabled\" onAction=\"RepurposeFileSave\"/></commands>");
            xml.Append(TabMarkup());
            xml.Append(ContextMarkup(application));
            return xml.Append("</customUI>").ToString();
        }

        internal static string TabMarkup()
        {
            return "<ribbon><tabs><tab id=\"OieTab\" label=\"Office in Editor\"><group id=\"OieChatGroup\" label=\"Chat context\">"
                + "<button id=\"OieRefresh\" label=\"Refresh chats\" onAction=\"RefreshChats\" getEnabled=\"GetRefreshEnabled\"/>"
                + "<dynamicMenu id=\"OieTargets\" getLabel=\"GetTargetLabel\" getContent=\"GetChatMenu\" invalidateContentOnDrop=\"true\"/>"
                + "<button id=\"OieSend\" label=\"Add to Chat\" onAction=\"AddToChat\" getEnabled=\"GetSendEnabled\"/>"
                + "<labelControl id=\"OieStatus\" getLabel=\"GetStatusLabel\"/>"
                + "</group><group id=\"OieSaveGroup\" label=\"本地保存\">"
                + "<toggleButton id=\"OieSavePeriodic\" label=\"定时保存\" getPressed=\"GetSavePeriodic\" onAction=\"SetSavePeriodic\"/>"
                + "<toggleButton id=\"OieSaveOnClose\" label=\"关标签保存\" getPressed=\"GetSaveOnClose\" onAction=\"SetSaveOnClose\"/>"
                + "<editBox id=\"OieSaveInterval\" label=\"间隔（秒）\" getText=\"GetSaveInterval\" onChange=\"SetSaveInterval\" maxLength=\"4\"/>"
                + "</group></tab></tabs></ribbon>";
        }

        internal static string ContextMarkup(string application)
        {
            var xml = new StringBuilder("<contextMenus>");
            if (Contexts.TryGetValue(application, out string[] contexts))
                for (int i = 0; i < contexts.Length; i++)
                    xml.Append("<contextMenu idMso=\"").Append(contexts[i]).Append("\"><button id=\"OieContext")
                        .Append(i).Append("\" label=\"Add to Chat\" insertBeforeMso=\"Cut\" onAction=\"AddToChat\" getEnabled=\"GetSendEnabled\"/></contextMenu>");
            return xml.Append("</contextMenus>").ToString();
        }
    }
}
