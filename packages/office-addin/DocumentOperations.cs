using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OfficeInEditor.OfficeAddin
{
    internal sealed class DocumentOperations
    {
        private const int MaxText = 8192;
        private const int MaxCells = 200;
        private readonly dynamic application;
        private readonly string kind;
        private readonly DocumentRegistry documents;

        internal DocumentOperations(object application, string kind, DocumentRegistry documents)
        {
            this.application = application;
            this.kind = kind;
            this.documents = documents;
        }

        internal object Dispatch(string method, object parameters)
        {
            try
            {
                switch (method)
                {
                    case "inspectDocument": return Inspect(parameters);
                    case "prepareEdit": return Prepare(parameters);
                    case "applyEdit": return Apply(parameters);
                    case "revertEdit": return Revert(parameters);
                    case "saveDocument": return Save(parameters);
                    default: throw new BridgeException("METHOD_NOT_FOUND");
                }
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("OFFICE_OPERATION_FAILED"); }
        }

        private object Inspect(object parameters)
        {
            string requested = Json.Text(parameters, "documentId");
            dynamic document = string.IsNullOrEmpty(requested) ? Active() : documents.Resolve(requested);
            var described = Describe(document);
            var list = new List<object>();
            foreach (object item in documents.OpenDocuments()) list.Add(Summary(item));
            described.Add("documents", list);
            described.Add("capabilities", Capabilities(document));
            return described;
        }

        private object Prepare(object parameters)
        {
            var operation = Operation(parameters);
            string name = Kind(operation);
            dynamic document = documents.Resolve(Json.Text(parameters, "documentId"));
            RequireWritable(document);
            if (!Supports(name)) throw new BridgeException("OPERATION_UNSUPPORTED");
            Dictionary<string, object> args = Args(operation);
            object recovery;
            object prepared;
            string before;
            switch (name)
            {
                case "word.insert_text":
                    string text = RequiredText(args);
                    int start = WordStart(document, args);
                    before = WordBeforeHash(document, start);
                    recovery = Json.Object("kind", name, "start", start, "text", text);
                    prepared = Json.Object("start", start, "length", text.Length);
                    break;
                case "excel.set_range_values":
                case "excel.set_range_formulas":
                    var previous = SheetRange(document, args);
                    if (name == "excel.set_range_formulas") FormulaMatrix(args, previous.Rows, previous.Columns);
                    else ValueMatrix(args, previous.Rows, previous.Columns);
                    before = Hash("excel", previous.Sheet, previous.Address, Join(previous.Cells));
                    recovery = Json.Object("kind", name, "sheet", previous.Sheet, "address", previous.Address, "cells", previous.Cells);
                    prepared = Json.Object("sheet", previous.Sheet, "address", previous.Address, "rows", previous.Rows, "columns", previous.Columns);
                    break;
                case "powerpoint.add_slide":
                    var ids = SlideIds(document);
                    before = Hash("slides", Join(ids));
                    recovery = Json.Object("kind", name, "previousSlideIds", ids);
                    prepared = Json.Object("index", ids.Count + 1);
                    break;
                case "powerpoint.add_textbox":
                    int slideId = RequiredInt(args, "slideId");
                    string boxText = RequiredText(args);
                    Bounds(args);
                    var shapes = ShapeIds(SlideById(document, slideId));
                    before = Hash("shapes", slideId.ToString(CultureInfo.InvariantCulture), Join(shapes));
                    recovery = Json.Object("kind", name, "slideId", slideId, "text", boxText, "previousShapeIds", shapes);
                    prepared = Json.Object("slideId", slideId, "length", boxText.Length);
                    break;
                default: throw new BridgeException("OPERATION_UNSUPPORTED");
            }
            return Json.Object("documentId", documents.Bind(document), "application", kind,
                "beforeHash", before, "recovery", recovery, "prepared", prepared);
        }

        private object Apply(object parameters)
        {
            var operation = Operation(parameters);
            string name = Kind(operation);
            dynamic document = documents.Resolve(Json.Text(parameters, "documentId"));
            RequireWritable(document);
            var prepared = Json.Map(Json.Get(parameters, "prepared")) ?? throw new BridgeException("INVALID_OPERATION");
            var recovery = Json.Map(Json.Get(parameters, "recovery")) ?? throw new BridgeException("RECOVERY_UNAVAILABLE");
            string before = Json.Text(parameters, "beforeHash");
            if (string.IsNullOrEmpty(before)) throw new BridgeException("INVALID_OPERATION");
            object result;
            string after;
            switch (name)
            {
                case "word.insert_text":
                    int start = Convert.ToInt32(Json.Get(recovery, "start"), CultureInfo.InvariantCulture);
                    string text = Convert.ToString(Json.Get(recovery, "text"), CultureInfo.InvariantCulture);
                    if (start != Convert.ToInt32(Json.Get(prepared, "start"), CultureInfo.InvariantCulture) || text != RequiredText(Args(operation)))
                        throw new BridgeException("INVALID_OPERATION");
                    if (WordBeforeHash(document, start) != before) throw new BridgeException("DOCUMENT_CONFLICT");
                    dynamic point = document.Range(start, start);
                    point.InsertBefore(text);
                    string landed = Slice(document, start, text.Length);
                    if (landed != text) throw new BridgeException("OFFICE_OPERATION_FAILED");
                    after = WordAfterHash(document, start, text);
                    result = Json.Object("inserted", true);
                    break;
                case "excel.set_range_values":
                case "excel.set_range_formulas":
                    var target = BindRange(document, prepared);
                    if (Hash("excel", target.Sheet, target.Address, Join(ReadCells(target.Range, target.Address).Cells)) != before)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    var encoded = name == "excel.set_range_formulas"
                        ? FormulaMatrix(Args(operation), target.Rows, target.Columns)
                        : ValueMatrix(Args(operation), target.Rows, target.Columns);
                    WriteCells(target.Range, encoded);
                    after = Hash("excel-after", target.Sheet, target.Address, Join(ReadCells(target.Range, target.Address).Cells));
                    result = Json.Object("updated", true);
                    break;
                case "powerpoint.add_slide":
                    var previous = ReadInts(Json.Get(recovery, "previousSlideIds"));
                    if (Hash("slides", Join(SlideIds(document))) != before) throw new BridgeException("DOCUMENT_CONFLICT");
                    int index = previous.Count + 1;
                    dynamic slide = document.Slides.Add(index, 12);
                    int slideId = Convert.ToInt32(slide.SlideID, CultureInfo.InvariantCulture);
                    after = Hash("slides-after", Join(SlideIds(document)), slideId.ToString(CultureInfo.InvariantCulture));
                    result = Json.Object("slideId", slideId);
                    break;
                case "powerpoint.add_textbox":
                    int boxSlide = Convert.ToInt32(Json.Get(recovery, "slideId"), CultureInfo.InvariantCulture);
                    string boxText = Convert.ToString(Json.Get(recovery, "text"), CultureInfo.InvariantCulture);
                    if (boxText != RequiredText(Args(operation))) throw new BridgeException("INVALID_OPERATION");
                    dynamic targetSlide = SlideById(document, boxSlide);
                    if (Hash("shapes", boxSlide.ToString(CultureInfo.InvariantCulture), Join(ShapeIds(targetSlide))) != before)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    var boxArgs = Args(operation);
                    dynamic shape = targetSlide.Shapes.AddTextbox(1, Number(boxArgs, "left"), Number(boxArgs, "top"),
                        Number(boxArgs, "width"), Number(boxArgs, "height"));
                    shape.TextFrame.TextRange.Text = boxText;
                    int shapeId = Convert.ToInt32(shape.Id, CultureInfo.InvariantCulture);
                    after = Hash("shapes-after", boxSlide.ToString(CultureInfo.InvariantCulture), Join(ShapeIds(targetSlide)),
                        shapeId.ToString(CultureInfo.InvariantCulture), boxText);
                    result = Json.Object("shapeId", shapeId);
                    break;
                default: throw new BridgeException("OPERATION_UNSUPPORTED");
            }
            return Json.Object("documentId", documents.Bind(document), "afterHash", after, "result", result);
        }

        private object Revert(object parameters)
        {
            dynamic document = documents.Resolve(Json.Text(parameters, "documentId"));
            RequireWritable(document);
            var recovery = Json.Map(Json.Get(parameters, "recovery")) ?? throw new BridgeException("RECOVERY_UNAVAILABLE");
            string name = Convert.ToString(Json.Get(recovery, "kind"), CultureInfo.InvariantCulture);
            string after = Json.Text(parameters, "afterHash");
            var outcome = Json.Map(Json.Get(parameters, "result"));
            switch (name)
            {
                case "word.insert_text":
                    int start = Convert.ToInt32(Json.Get(recovery, "start"), CultureInfo.InvariantCulture);
                    string text = Convert.ToString(Json.Get(recovery, "text"), CultureInfo.InvariantCulture);
                    if (WordAfterHash(document, start, text) != after || Slice(document, start, text.Length) != text)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    dynamic range = document.Range(start, start + text.Length);
                    range.Text = "";
                    break;
                case "excel.set_range_values":
                case "excel.set_range_formulas":
                    var target = BindRange(document, Json.Object("sheet", Json.Get(recovery, "sheet"), "address", Json.Get(recovery, "address"),
                        "rows", RowsOf(Json.Get(recovery, "cells")), "columns", ColumnsOf(Json.Get(recovery, "cells"))));
                    if (Hash("excel-after", target.Sheet, target.Address, Join(ReadCells(target.Range, target.Address).Cells)) != after)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    WriteCells(target.Range, Matrix(Json.Get(recovery, "cells"), target.Rows, target.Columns));
                    break;
                case "powerpoint.add_slide":
                    int slideId = Convert.ToInt32(Json.Get(outcome, "slideId"), CultureInfo.InvariantCulture);
                    var previous = ReadInts(Json.Get(recovery, "previousSlideIds"));
                    if (previous.Contains(slideId) || Hash("slides-after", Join(SlideIds(document)), slideId.ToString(CultureInfo.InvariantCulture)) != after)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    SlideById(document, slideId).Delete();
                    break;
                case "powerpoint.add_textbox":
                    int boxSlide = Convert.ToInt32(Json.Get(recovery, "slideId"), CultureInfo.InvariantCulture);
                    int shapeId = Convert.ToInt32(Json.Get(outcome, "shapeId"), CultureInfo.InvariantCulture);
                    string boxText = Convert.ToString(Json.Get(recovery, "text"), CultureInfo.InvariantCulture);
                    dynamic slide = SlideById(document, boxSlide);
                    dynamic shape = ShapeById(slide, shapeId);
                    string current = Convert.ToString(shape.TextFrame.TextRange.Text, CultureInfo.InvariantCulture);
                    if (current != boxText || Hash("shapes-after", boxSlide.ToString(CultureInfo.InvariantCulture), Join(ShapeIds(slide)),
                        shapeId.ToString(CultureInfo.InvariantCulture), boxText) != after)
                        throw new BridgeException("DOCUMENT_CONFLICT");
                    shape.Delete();
                    break;
                default: throw new BridgeException("OPERATION_UNSUPPORTED");
            }
            return Json.Object("documentId", documents.Bind(document), "reverted", true);
        }

        private object Save(object parameters)
        {
            dynamic document = string.IsNullOrEmpty(Json.Text(parameters, "documentId"))
                ? Active() : documents.Resolve(Json.Text(parameters, "documentId"));
            if (!CanManualSave(document)) throw new BridgeException("MANUAL_SAVE_UNAVAILABLE");
            document.Save();
            return Json.Object("documentId", documents.Bind(document), "saved", true, "intent", "manual_save");
        }

        private object Active()
        {
            dynamic document;
            switch (kind)
            {
                case "word":
                    if (Convert.ToInt32(application.Documents.Count, CultureInfo.InvariantCulture) == 0) throw new BridgeException("NO_DOCUMENT");
                    document = application.ActiveDocument; break;
                case "excel":
                    if (Convert.ToInt32(application.Workbooks.Count, CultureInfo.InvariantCulture) == 0) throw new BridgeException("NO_DOCUMENT");
                    document = application.ActiveWorkbook; break;
                case "powerpoint":
                    if (Convert.ToInt32(application.Presentations.Count, CultureInfo.InvariantCulture) == 0) throw new BridgeException("NO_DOCUMENT");
                    document = application.ActivePresentation; break;
                default: throw new BridgeException("UNSUPPORTED_APPLICATION");
            }
            documents.Bind(document);
            return document;
        }

        private Dictionary<string, object> Describe(dynamic document)
        {
            string name = Convert.ToString(document.Name, CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(name) || name.Length > 256) throw new BridgeException("INVALID_DOCUMENT");
            bool dirty = true;
            string state;
            try { dirty = !Convert.ToBoolean(document.Saved, CultureInfo.InvariantCulture); } catch { }
            try { state = Convert.ToBoolean(document.AutoSaveOn, CultureInfo.InvariantCulture) ? "enabled" : "off_or_unavailable"; }
            catch { state = "unsupported"; }
            return Json.Object("application", kind, "documentId", documents.Bind(document), "documentName", name,
                "dirty", dirty, "autoSaveState", state, "nativeAutoSaveLocalRedirect", false, "nativeEmbedding", false);
        }

        private Dictionary<string, object> Summary(dynamic document)
        {
            var described = Describe(document);
            described.Remove("dirty");
            return described;
        }

        private Dictionary<string, object> Capabilities(dynamic document)
        {
            var kinds = new List<string>();
            if (!IsReadOnly(document))
            {
                if (kind == "word") kinds.Add("word.insert_text");
                if (kind == "excel") { kinds.Add("excel.set_range_values"); kinds.Add("excel.set_range_formulas"); }
                if (kind == "powerpoint") { kinds.Add("powerpoint.add_slide"); kinds.Add("powerpoint.add_textbox"); }
            }
            return Json.Object("writeKinds", kinds.ToArray(), "manualSave", CanManualSave(document), "revert", !IsReadOnly(document),
                "nativeEmbedding", false, "nativeAutoSaveLocalRedirect", false);
        }

        private bool Supports(string name)
        {
            switch (kind)
            {
                case "word": return name == "word.insert_text";
                case "excel": return name == "excel.set_range_values" || name == "excel.set_range_formulas";
                case "powerpoint": return name == "powerpoint.add_slide" || name == "powerpoint.add_textbox";
                default: return false;
            }
        }

        private static Dictionary<string, object> Operation(object parameters) =>
            Json.Map(Json.Get(parameters, "operation")) ?? throw new BridgeException("INVALID_OPERATION");
        private static string Kind(object operation)
        {
            string name = Json.Text(operation, "kind");
            if (string.IsNullOrEmpty(name) || name.Length > 64) throw new BridgeException("INVALID_OPERATION");
            return name;
        }
        private static Dictionary<string, object> Args(object operation) =>
            Json.Map(Json.Get(operation, "args")) ?? throw new BridgeException("INVALID_OPERATION");

        private static string RequiredText(object args)
        {
            string text = Json.Text(args, "text");
            if (string.IsNullOrEmpty(text) || text.Length > MaxText) throw new BridgeException("TEXT_TOO_LARGE");
            foreach (char c in text)
                if (char.IsControl(c) && c != '\t' && c != '\r' && c != '\n') throw new BridgeException("INVALID_OPERATION");
            return text;
        }

        private static int WordStart(dynamic document, object args)
        {
            int end = Convert.ToInt32(document.Content.End, CultureInfo.InvariantCulture);
            object supplied = Json.Get(args, "offset");
            int start = supplied == null ? Math.Max(0, end - 1) : Convert.ToInt32(supplied, CultureInfo.InvariantCulture);
            if (start < 0 || (end > 0 && start >= end) || (end <= 0 && start != 0)) throw new BridgeException("INVALID_RANGE");
            return start;
        }

        private static string Slice(dynamic document, int start, int length)
        {
            if (length < 0 || length > MaxText) throw new BridgeException("INVALID_RANGE");
            if (length == 0) return "";
            string text = Convert.ToString(document.Range(start, start + length).Text, CultureInfo.InvariantCulture) ?? "";
            return text;
        }

        private static string WordBeforeHash(dynamic document, int start)
        {
            int end = Convert.ToInt32(document.Content.End, CultureInfo.InvariantCulture);
            string left = start <= 0 || end <= 0 ? "" : Slice(document, Math.Max(0, start - 24), Math.Min(24, start));
            string right = end <= start ? "" : Slice(document, start, Math.Min(24, end - start));
            return Hash("word", start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), left, right);
        }

        private static string WordAfterHash(dynamic document, int start, string text)
        {
            int end = Convert.ToInt32(document.Content.End, CultureInfo.InvariantCulture);
            return Hash("word-after", start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture),
                Slice(document, start, text.Length));
        }

        private sealed class CellBlock
        {
            internal string Sheet, Address;
            internal int Rows, Columns;
            internal List<List<string>> Cells;
            internal dynamic Range;
        }

        private CellBlock SheetRange(dynamic workbook, object args)
        {
            dynamic sheet = Sheet(workbook, args);
            return Capture(sheet, Address(args));
        }

        private static dynamic Sheet(dynamic workbook, object args)
        {
            string name = Json.Text(args, "sheet");
            try { return string.IsNullOrEmpty(name) ? workbook.ActiveSheet : workbook.Worksheets.Item(name); }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("INVALID_RANGE"); }
        }

        private static string Address(object args)
        {
            string address = Json.Text(args, "address");
            if (string.IsNullOrEmpty(address) || address.Length > 32 || address.IndexOf('!') >= 0 || address.IndexOf(',') >= 0)
                throw new BridgeException("INVALID_RANGE");
            return address;
        }

        private CellBlock BindRange(dynamic workbook, object prepared)
        {
            string sheet = Convert.ToString(Json.Get(prepared, "sheet"), CultureInfo.InvariantCulture);
            string address = Convert.ToString(Json.Get(prepared, "address"), CultureInfo.InvariantCulture);
            dynamic worksheet;
            try { worksheet = string.IsNullOrEmpty(sheet) ? workbook.ActiveSheet : workbook.Worksheets.Item(sheet); }
            catch { throw new BridgeException("INVALID_RANGE"); }
            var block = Capture(worksheet, address);
            if (block.Rows != Convert.ToInt32(Json.Get(prepared, "rows"), CultureInfo.InvariantCulture) ||
                block.Columns != Convert.ToInt32(Json.Get(prepared, "columns"), CultureInfo.InvariantCulture))
                throw new BridgeException("DOCUMENT_CONFLICT");
            return block;
        }

        private static CellBlock Capture(dynamic sheet, string address)
        {
            dynamic range;
            try { range = sheet.Range(address); }
            catch { throw new BridgeException("INVALID_RANGE"); }
            return ReadCells(range, address);
        }

        private static CellBlock ReadCells(dynamic range, string address)
        {
            int rows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            int columns = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            if (rows < 1 || columns < 1 || (long)rows * columns > MaxCells) throw new BridgeException("SELECTION_TOO_LARGE");
            string sheet = Convert.ToString(range.Worksheet.Name, CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(sheet) || sheet.Length > 31 || string.IsNullOrEmpty(address)) throw new BridgeException("INVALID_RANGE");
            try { if (Convert.ToBoolean(range.Worksheet.ProtectContents, CultureInfo.InvariantCulture)) throw new BridgeException("DOCUMENT_PROTECTED"); }
            catch (BridgeException) { throw; }
            catch { }
            var cells = new List<List<string>>();
            for (int row = 1; row <= rows; row++)
            {
                var line = new List<string>();
                for (int column = 1; column <= columns; column++) line.Add(Encode(range.Cells.Item(row, column)));
                cells.Add(line);
            }
            return new CellBlock { Sheet = sheet, Address = address, Rows = rows, Columns = columns, Cells = cells, Range = range };
        }

        private static string Encode(dynamic cell)
        {
            object formula = cell.Formula;
            if (formula is string formulaText && formulaText.StartsWith("=", StringComparison.Ordinal))
            {
                if (formulaText.Length > 1024) throw new BridgeException("TEXT_TOO_LARGE");
                return "f:" + formulaText;
            }
            object value = cell.Value2;
            if (value == null) return "z:";
            if (value is bool bit) return "b:" + (bit ? "true" : "false");
            if (value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal)
                return "n:" + Convert.ToString(value, CultureInfo.InvariantCulture);
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            if (text.Length > MaxText) throw new BridgeException("TEXT_TOO_LARGE");
            return "s:" + text;
        }

        private static List<List<string>> ValueMatrix(object args, int rows, int columns)
        {
            var matrix = new List<List<string>>();
            foreach (object row in Json.Array(Json.Get(args, "values")))
            {
                var line = new List<string>();
                foreach (object cell in Json.Array(row)) line.Add(EncodeUser(cell));
                if (line.Count != columns) throw new BridgeException("INVALID_RANGE");
                matrix.Add(line);
            }
            if (matrix.Count != rows) throw new BridgeException("INVALID_RANGE");
            return matrix;
        }

        private static string EncodeUser(object cell)
        {
            if (cell == null) return "z:";
            if (cell is bool bit) return "b:" + (bit ? "true" : "false");
            if (cell is byte || cell is short || cell is int || cell is long || cell is float || cell is double || cell is decimal)
                return "n:" + Convert.ToString(cell, CultureInfo.InvariantCulture);
            string text = Convert.ToString(cell, CultureInfo.InvariantCulture) ?? "";
            if (text.StartsWith("=", StringComparison.Ordinal)) throw new BridgeException("INVALID_OPERATION");
            if (text.Length > MaxText) throw new BridgeException("TEXT_TOO_LARGE");
            return text.Length == 0 ? "z:" : "s:" + text;
        }

        private static List<List<string>> FormulaMatrix(object args, int rows, int columns)
        {
            var matrix = Matrix(Json.Get(args, "formulas"), rows, columns);
            for (int row = 0; row < matrix.Count; row++)
                for (int column = 0; column < matrix[row].Count; column++)
                {
                    string text = matrix[row][column] ?? "";
                    if (!text.StartsWith("=", StringComparison.Ordinal) || text.Length > 1024) throw new BridgeException("INVALID_OPERATION");
                    matrix[row][column] = "f:" + text;
                }
            return matrix;
        }

        private static void WriteCells(dynamic range, List<List<string>> cells)
        {
            for (int row = 0; row < cells.Count; row++)
                for (int column = 0; column < cells[row].Count; column++)
                {
                    dynamic cell = range.Cells.Item(row + 1, column + 1);
                    string encoded = cells[row][column];
                    if (encoded.StartsWith("f:", StringComparison.Ordinal)) cell.Formula = encoded.Substring(2);
                    else if (encoded.StartsWith("b:", StringComparison.Ordinal)) cell.Value2 = encoded.EndsWith("true", StringComparison.Ordinal);
                    else if (encoded.StartsWith("n:", StringComparison.Ordinal)) cell.Value2 = double.Parse(encoded.Substring(2), CultureInfo.InvariantCulture);
                    else if (encoded.StartsWith("z:", StringComparison.Ordinal)) cell.Value2 = null;
                    else cell.Value2 = encoded.Substring(2);
                }
        }

        private static List<int> SlideIds(dynamic presentation)
        {
            int count = Convert.ToInt32(presentation.Slides.Count, CultureInfo.InvariantCulture);
            if (count > 256) throw new BridgeException("SELECTION_TOO_LARGE");
            var ids = new List<int>();
            for (int i = 1; i <= count; i++) ids.Add(Convert.ToInt32(presentation.Slides.Item(i).SlideID, CultureInfo.InvariantCulture));
            return ids;
        }

        private static dynamic SlideById(dynamic presentation, int slideId)
        {
            int count = Convert.ToInt32(presentation.Slides.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                dynamic slide = presentation.Slides.Item(i);
                if (Convert.ToInt32(slide.SlideID, CultureInfo.InvariantCulture) == slideId) return slide;
            }
            throw new BridgeException("DOCUMENT_CONFLICT");
        }

        private static List<int> ShapeIds(dynamic slide)
        {
            int count = Convert.ToInt32(slide.Shapes.Count, CultureInfo.InvariantCulture);
            if (count > 200) throw new BridgeException("SELECTION_TOO_LARGE");
            var ids = new List<int>();
            for (int i = 1; i <= count; i++) ids.Add(Convert.ToInt32(slide.Shapes.Item(i).Id, CultureInfo.InvariantCulture));
            return ids;
        }

        private static dynamic ShapeById(dynamic slide, int shapeId)
        {
            int count = Convert.ToInt32(slide.Shapes.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                dynamic shape = slide.Shapes.Item(i);
                if (Convert.ToInt32(shape.Id, CultureInfo.InvariantCulture) == shapeId) return shape;
            }
            throw new BridgeException("DOCUMENT_CONFLICT");
        }

        private static void Bounds(object args)
        {
            foreach (string name in new[] { "left", "top", "width", "height" })
            {
                double value = Number(args, name);
                if (name == "width" || name == "height") { if (value < 1 || value > 20000) throw new BridgeException("INVALID_RANGE"); }
                else if (value < 0 || value > 20000) throw new BridgeException("INVALID_RANGE");
            }
        }

        private static double Number(object args, string name)
        {
            object value = Json.Get(args, name);
            double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(number) || double.IsInfinity(number)) throw new BridgeException("INVALID_RANGE");
            return number;
        }

        private static int RequiredInt(object args, string name) => Convert.ToInt32(Json.Get(args, name), CultureInfo.InvariantCulture);

        private static void RequireWritable(dynamic document)
        {
            if (IsReadOnly(document)) throw new BridgeException("DOCUMENT_PROTECTED");
            try
            {
                object protection = document.ProtectionType;
                if (protection != null && Convert.ToInt32(protection, CultureInfo.InvariantCulture) != -1)
                    throw new BridgeException("DOCUMENT_PROTECTED");
            }
            catch (BridgeException) { throw; }
            catch { }
        }

        private static bool IsReadOnly(dynamic document)
        {
            try { return Convert.ToBoolean(document.ReadOnly, CultureInfo.InvariantCulture); }
            catch { return false; }
        }

        private static bool CanManualSave(dynamic document)
        {
            if (IsReadOnly(document)) return false;
            string fullName;
            try { fullName = Convert.ToString(document.FullName, CultureInfo.InvariantCulture); }
            catch { return false; }
            if (string.IsNullOrEmpty(fullName) || fullName.Length < 4 || fullName.Length > 240)
                return false;
            if (!char.IsLetter(fullName[0]) || fullName[1] != ':' || (fullName[2] != '\\' && fullName[2] != '/'))
                return false;
            string extension = Path.GetExtension(fullName);
            return extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase);
        }

        private static List<List<string>> Matrix(object value, int rows, int columns)
        {
            var matrix = new List<List<string>>();
            foreach (object row in Json.Array(value))
            {
                var line = new List<string>();
                foreach (object cell in Json.Array(row))
                {
                    string text = Convert.ToString(cell, CultureInfo.InvariantCulture) ?? "";
                    if (text.Length > MaxText) throw new BridgeException("TEXT_TOO_LARGE");
                    line.Add(text);
                }
                if (line.Count != columns) throw new BridgeException("INVALID_RANGE");
                matrix.Add(line);
            }
            if (matrix.Count != rows) throw new BridgeException("INVALID_RANGE");
            return matrix;
        }

        private static int RowsOf(object value)
        {
            int count = 0;
            foreach (object unused in Json.Array(value)) count++;
            return count;
        }

        private static int ColumnsOf(object value)
        {
            foreach (object row in Json.Array(value))
            {
                int count = 0;
                foreach (object unused in Json.Array(row)) count++;
                return count;
            }
            return 0;
        }

        private static List<int> ReadInts(object value)
        {
            var result = new List<int>();
            foreach (object item in Json.Array(value)) result.Add(Convert.ToInt32(item, CultureInfo.InvariantCulture));
            return result;
        }

        private static string Join(IEnumerable values)
        {
            var builder = new StringBuilder();
            foreach (object value in values)
            {
                if (builder.Length > 0) builder.Append('\n');
                if (value is IEnumerable nested && !(value is string)) builder.Append(Join(nested));
                else builder.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }

        private static string Hash(params string[] parts)
        {
            using (var sha = SHA256.Create())
            {
                byte[] bytes = Wire.Utf8.GetBytes(string.Join("\n", parts));
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
