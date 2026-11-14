using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeInEditor.OfficeAddin
{
    internal sealed class SelectionCapture
    {
        internal const int MaxSelectionBytes = 1024 * 1024;
        private const int MaxCells = 5000;
        private const int MaxShapes = 100;
        private readonly dynamic application;
        private readonly string kind;
        private readonly string officeId;
        private Stopwatch captureClock;
        internal DocumentRegistry Documents { get; }

        internal SelectionCapture(object application, string kind, string officeId)
        { this.application = application; this.kind = kind; this.officeId = officeId;
            Documents = new DocumentRegistry(application, kind, officeId); }

        internal object Capture()
        {
            captureClock = Stopwatch.StartNew();
            Dictionary<string, object> result;
            try
            {
                switch (kind)
                {
                    case "word": result = Word(); break;
                    case "excel": result = Excel(); break;
                    case "powerpoint": result = PowerPoint(); break;
                    default: throw new BridgeException("UNSUPPORTED_APPLICATION");
                }
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("NO_SUPPORTED_SELECTION"); }
            if (Wire.Utf8.GetByteCount(Json.Encode(result)) > MaxSelectionBytes)
                throw new BridgeException("SELECTION_TOO_LARGE");
            return result;
        }

        private Dictionary<string, object> Metadata(object documentObject, string selectionKind,
            string text, object locator)
        {
            var result = Documents.Describe(documentObject);
            result.Add("kind", selectionKind); result.Add("text", text); result.Add("locator", locator);
            return result;
        }

        private Dictionary<string, object> Word()
        {
            if (Convert.ToInt32(application.Documents.Count) == 0) throw new BridgeException("NO_DOCUMENT");
            dynamic document = application.ActiveDocument;
            dynamic range = application.Selection.Range;
            int start = Convert.ToInt32(range.Start), end = Convert.ToInt32(range.End);
            if (end <= start) throw new BridgeException("NO_SELECTION");
            // Guard before requesting a potentially huge COM BSTR. No silently truncated text.
            if ((long)end - start > MaxSelectionBytes / 2) throw new BridgeException("SELECTION_TOO_LARGE");
            string text = Convert.ToString(range.Text, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text) || text.Replace("\u0001", "").Length == 0)
                throw new BridgeException("NO_SUPPORTED_SELECTION");
            return Metadata(document, "word-range", text, Json.Object("storyType", Convert.ToInt32(range.StoryType),
                "start", start, "end", end));
        }

        private Dictionary<string, object> Excel()
        {
            if (Convert.ToInt32(application.Workbooks.Count) == 0) throw new BridgeException("NO_DOCUMENT");
            dynamic workbook = application.ActiveWorkbook;
            dynamic selected = application.Selection;
            // Only worksheet Range selections are supported; no ActiveCell fallback for objects.
            int areaCount = Convert.ToInt32(selected.Areas.Count);
            if (areaCount < 1 || areaCount > 32) throw new BridgeException("SELECTION_TOO_LARGE");
            if (Convert.ToDouble(selected.CountLarge, CultureInfo.InvariantCulture) > MaxCells)
                throw new BridgeException("SELECTION_TOO_LARGE");
            dynamic worksheet = selected.Worksheet;
            var areas = new List<object>();
            var text = new StringBuilder();
            long capturedCharacters = 0;
            for (int i = 1; i <= areaCount; i++)
            {
                dynamic area = selected.Areas.Item(i);
                int rows = Convert.ToInt32(area.Rows.Count), columns = Convert.ToInt32(area.Columns.Count);
                if ((long)rows * columns > MaxCells) throw new BridgeException("SELECTION_TOO_LARGE");
                string address = Convert.ToString(area.Address[true, true, 1, false], CultureInfo.InvariantCulture);
                var valueRows = new List<object>();
                var formulaRows = new List<object>();
                if (i > 1) text.AppendLine();
                text.Append(address).AppendLine();
                for (int row = 1; row <= rows; row++)
                {
                    var rowValues = new List<object>();
                    var rowFormulas = new List<object>();
                    for (int col = 1; col <= columns; col++)
                    {
                        CheckTime();
                        // Read one cell at a time: a large range of long strings must not allocate
                        // an unbounded COM array before the reference-size guard gets a chance to run.
                        dynamic selectedCell = area.Cells.Item(row, col);
                        object cell = selectedCell.Value2;
                        object formula = selectedCell.Formula;
                        string valueText = Scalar(cell);
                        string formulaText = formula as string;
                        capturedCharacters += valueText.Length + (formulaText?.Length ?? 0);
                        if (capturedCharacters > MaxSelectionBytes / 2)
                            throw new BridgeException("SELECTION_TOO_LARGE");
                        rowValues.Add(cell == null ? null : (object)valueText);
                        rowFormulas.Add(formulaText != null && formulaText.StartsWith("=", StringComparison.Ordinal)
                            ? formulaText : null);
                        if (col > 1) text.Append('\t');
                        text.Append(valueText.Replace("\t", " ").Replace("\r", " ").Replace("\n", " "));
                        CheckText(text);
                    }
                    text.AppendLine(); valueRows.Add(rowValues); formulaRows.Add(rowFormulas);
                }
                areas.Add(Json.Object("address", address, "values", valueRows, "formulas", formulaRows));
            }
            return Metadata(workbook, "excel-range", text.ToString(), Json.Object(
                "worksheet", Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture),
                "areas", areas));
        }
        private static string Scalar(object source)
        {
            if (source == null) return "";
            if (source is string s)
            { if (s.Length > MaxSelectionBytes / 2) throw new BridgeException("SELECTION_TOO_LARGE"); return s; }
            if (source is IConvertible) return Convert.ToString(source, CultureInfo.InvariantCulture);
            return "[cell error or unsupported value]";
        }
        private static void CheckText(StringBuilder text)
        { if (text.Length > MaxSelectionBytes / 2) throw new BridgeException("SELECTION_TOO_LARGE"); }
        private void CheckTime()
        { if (captureClock.ElapsedMilliseconds > 2500) throw new BridgeException("SELECTION_CAPTURE_BUDGET_EXCEEDED"); }

        private Dictionary<string, object> PowerPoint()
        {
            if (Convert.ToInt32(application.Presentations.Count) == 0) throw new BridgeException("NO_DOCUMENT");
            dynamic presentation = application.ActivePresentation;
            dynamic selected = application.ActiveWindow.Selection;
            int selectionType = Convert.ToInt32(selected.Type);
            var text = new StringBuilder();
            var slides = new List<object>();
            if (selectionType == 3) // ppSelectionText
            {
                dynamic selectedText = selected.TextRange;
                if (Convert.ToInt32(selectedText.Length) > MaxSelectionBytes / 2)
                    throw new BridgeException("SELECTION_TOO_LARGE");
                string value = Convert.ToString(selectedText.Text, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(value)) throw new BridgeException("NO_SELECTION");
                text.Append(value);
                var ids = new List<object>();
                int shapeCount = 0;
                try { shapeCount = Convert.ToInt32(selected.ShapeRange.Count); } catch { }
                if (shapeCount > MaxShapes) throw new BridgeException("SELECTION_TOO_LARGE");
                for (int i = 1; i <= shapeCount; i++)
                { CheckTime(); ids.Add(Convert.ToInt32(selected.ShapeRange.Item(i).Id)); }
                dynamic slide = application.ActiveWindow.View.Slide;
                slides.Add(Json.Object("slideId", Convert.ToInt32(slide.SlideID), "shapeIds", ids,
                    "textStart", Convert.ToInt32(selectedText.Start), "textLength", Convert.ToInt32(selectedText.Length)));
            }
            else if (selectionType == 2) // ppSelectionShapes
            {
                dynamic slide = application.ActiveWindow.View.Slide;
                slides.Add(Json.Object("slideId", Convert.ToInt32(slide.SlideID),
                    "shapes", ShapeContent(selected.ShapeRange, text)));
            }
            else if (selectionType == 1) // ppSelectionSlides
            {
                int count = Convert.ToInt32(selected.SlideRange.Count);
                if (count > 10) throw new BridgeException("SELECTION_TOO_LARGE");
                int totalShapes = 0;
                for (int i = 1; i <= count; i++)
                {
                    CheckTime();
                    dynamic slide = selected.SlideRange.Item(i);
                    totalShapes += Convert.ToInt32(slide.Shapes.Count);
                    if (totalShapes > MaxShapes) throw new BridgeException("SELECTION_TOO_LARGE");
                    text.Append("Slide ").Append(Convert.ToInt32(slide.SlideIndex)).AppendLine();
                    slides.Add(Json.Object("slideId", Convert.ToInt32(slide.SlideID),
                        "shapes", ShapeContent(slide.Shapes, text)));
                }
            }
            else throw new BridgeException("NO_SUPPORTED_SELECTION");
            CheckText(text);
            return Metadata(presentation, "powerpoint-selection", text.ToString(),
                Json.Object("selectionType", selectionType, "slides", slides));
        }

        private List<object> ShapeContent(dynamic shapes, StringBuilder text)
        {
            int count = Convert.ToInt32(shapes.Count);
            if (count > MaxShapes) throw new BridgeException("SELECTION_TOO_LARGE");
            var result = new List<object>();
            for (int i = 1; i <= count; i++)
            {
                CheckTime();
                dynamic shape = shapes.Item(i);
                string value = "";
                if (Convert.ToInt32(shape.HasTextFrame) != 0)
                {
                    dynamic range = shape.TextFrame.TextRange;
                    if (Convert.ToInt32(range.Length) > MaxSelectionBytes / 2)
                        throw new BridgeException("SELECTION_TOO_LARGE");
                    value = Convert.ToString(range.Text, CultureInfo.InvariantCulture);
                }
                string name = Convert.ToString(shape.Name, CultureInfo.InvariantCulture);
                if (name != null && name.Length > 512) throw new BridgeException("SELECTION_TOO_LARGE");
                result.Add(Json.Object("shapeId", Convert.ToInt32(shape.Id), "name", name,
                    "shapeType", Convert.ToInt32(shape.Type), "text", value));
                text.Append(name).Append(": ").AppendLine(string.IsNullOrEmpty(value) ? "[non-text shape]" : value);
                CheckText(text);
            }
            return result;
        }
    }
}
