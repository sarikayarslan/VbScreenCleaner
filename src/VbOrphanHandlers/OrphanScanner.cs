// Finds event-handler procedures (xxx_Click, xxx_Change, ...) in VB5/VB6 forms whose
// control does not exist in the form designer, plus control-array Index mismatches.
// Read-only: VB files are never written.
using System.Text;
using System.Text.RegularExpressions;

namespace VbOrphanHandlers;

public static class OrphanScanner
{
    static readonly string[] FormExts = [".frm", ".ctl", ".pag", ".dob"];
    static readonly string[] SelfPrefixes = ["Form", "MDIForm", "UserControl", "PropertyPage", "UserDocument"];
    static readonly HashSet<string> KnownEvents = new(StringComparer.OrdinalIgnoreCase)
    {
        "Click", "DblClick", "Change", "GotFocus", "LostFocus", "KeyDown", "KeyUp", "KeyPress",
        "MouseDown", "MouseUp", "MouseMove", "Validate", "Scroll", "Timer", "Resize", "Paint",
        "DragDrop", "DragOver", "OLEDragDrop", "OLEDragOver", "OLECompleteDrag", "OLEGiveFeedback",
        "OLESetData", "OLEStartDrag", "DropDown", "PathChange", "PatternChange", "ItemCheck",
        "LinkClose", "LinkError", "LinkExecute", "LinkNotify", "LinkOpen", "LinkChange",
        "Load", "Unload", "QueryUnload", "Activate", "Deactivate", "Initialize", "Terminate",
        "ButtonClick", "ButtonMenuClick", "ButtonDropDown", "ItemClick", "NodeClick", "ColumnClick",
        "Expand", "Collapse", "BeforeLabelEdit", "AfterLabelEdit", "RowColChange", "EnterCell",
        "LeaveCell", "SelChange", "Compare", "BeforeUpdate", "AfterUpdate", "BeforeColUpdate",
        "AfterColUpdate", "BeforeInsert", "AfterInsert", "BeforeDelete", "AfterDelete", "HeadClick",
        "DateClick", "DateDblClick", "Reposition", "Error", "WillMove", "MoveComplete",
        "StateChanged", "DataArrival", "Connect", "ConnectionRequest", "Close", "SendComplete",
        "SendProgress", "Updated", "Format", "Unformat", "Validation", "ObjectEvent", "Hide", "Show",
        "Changed", "SelectionChanged", "TabClick", "BeforeClick", "ClickEvent", "AfterColEdit",
        "BeforeColEdit", "ColEdit", "ButtonUp", "ButtonDown", "SpinUp", "SpinDown", "UpClick", "DownClick",
    };
    static readonly Regex ProcRx = new(@"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+(\w+)\s*\((.*?)\)", RegexOptions.IgnoreCase);

    static Encoding VbEncoding
    {
        get
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(1254);
        }
    }

    /// <summary>inputs: project folders, .vbp files or extra (lib) folders. Returns the report text.</summary>
    public static string Scan(IReadOnlyList<string> inputs)
    {
        var enc = VbEncoding;
        var warnings = new List<string>();
        var files = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var input in inputs)
        {
            if (Directory.Exists(input))
            {
                foreach (var f in Directory.EnumerateFiles(input, "*.*", SearchOption.AllDirectories))
                    if (FormExts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)) files.Add(Path.GetFullPath(f));
            }
            else if (File.Exists(input) && input.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase))
            {
                var baseDir = Path.GetDirectoryName(Path.GetFullPath(input))!;
                foreach (var line in File.ReadAllLines(input, enc))
                {
                    var m = Regex.Match(line, @"^(Form|UserControl|PropertyPage|UserDocument)=(.+)$", RegexOptions.IgnoreCase);
                    if (!m.Success) continue;
                    var rel = m.Groups[2].Value.Trim();
                    if (rel.Contains(';')) rel = rel[(rel.IndexOf(';') + 1)..].Trim();
                    var full = Path.GetFullPath(Path.Combine(baseDir, rel));
                    if (File.Exists(full)) files.Add(full);
                    else warnings.Add($".vbp içindeki dosya bulunamadı: {full}");
                }
            }
            else throw new FileNotFoundException($"Bulunamadı: {input}");
        }

        var root = inputs.Select(p => Directory.Exists(p) ? Path.GetFullPath(p) : Path.GetDirectoryName(Path.GetFullPath(p))!).First();
        var orphans = new List<string>();
        var possible = new List<string>();
        var indexIssues = new List<string>();
        var indexUsage = new List<string>();
        int handlerCount = 0;

        foreach (var file in files)
            ScanFile(file, Path.GetRelativePath(root, file), enc, orphans, indexIssues, indexUsage, possible, ref handlerCount);

        var report = new StringBuilder();
        report.AppendLine("VB5 Sahipsiz Olay Prosedürleri Raporu");
        report.AppendLine($"Kaynak: {string.Join("; ", inputs)}");
        report.AppendLine($"Taranan form dosyası: {files.Count}, olay prosedürü: {handlerCount}");
        foreach (var w in warnings) report.AppendLine($"Uyarı: {w}");
        report.AppendLine();
        void Section(string title, List<string> items)
        {
            report.AppendLine($"== {title} ({items.Count}) ==");
            if (items.Count == 0) report.AppendLine("(yok)");
            foreach (var s in items) report.AppendLine(s);
            report.AppendLine();
        }
        Section("1. Kontrolü olmayan olay prosedürleri", orphans);
        Section("2. Index uyuşmazlıkları (kontrol dizisi / tekil kontrol)", indexIssues);
        Section("3. Kontrol dizilerinde index kullanımı (Select Case Index / If Index = n)", indexUsage);
        Section("4. Kontrol olmayan nesne_Ad prosedürleri (bilinmeyen olay adı, elle kontrol edin)", possible);
        report.AppendLine("Not: Form_/MDIForm_/UserControl_ olayları, WithEvents değişkenleri ve Implements arayüz metotları raporlanmaz.");
        return report.ToString();
    }

    // Compares Index values used in the handler body (Select Case Index / If Index = n) with the
    // designer's Index values. Returns report lines, or null when there is nothing to report.
    static string? AnalyzeIndexUsage(List<(int Line, string Text)> logical, int headerAt, SortedSet<int> defined)
    {
        var used = new SortedDictionary<int, int>();   // index -> first line
        var manual = new List<string>();
        var selectStack = new Stack<bool>();           // true = Select Case Index
        bool caseElse = false;
        for (int i = headerAt + 1; i < logical.Count; i++)
        {
            var (lineNo, raw) = logical[i];
            foreach (var stmt in SplitStatements(StripComment(raw)))
            {
                var s = stmt.Trim();
                if (Regex.IsMatch(s, @"^End\s+Sub\b", RegexOptions.IgnoreCase)) goto done;
                if (Regex.IsMatch(s, @"^Select\s+Case\s", RegexOptions.IgnoreCase))
                {
                    selectStack.Push(Regex.IsMatch(s, @"^Select\s+Case\s+Index\s*$", RegexOptions.IgnoreCase));
                    continue;
                }
                if (Regex.IsMatch(s, @"^End\s+Select\b", RegexOptions.IgnoreCase)) { if (selectStack.Count > 0) selectStack.Pop(); continue; }
                var c = Regex.Match(s, @"^Case\s+(.+)$", RegexOptions.IgnoreCase);
                if (c.Success && selectStack.Count > 0 && selectStack.Peek())
                {
                    if (Regex.IsMatch(c.Groups[1].Value, @"^Else\b", RegexOptions.IgnoreCase)) { caseElse = true; continue; }
                    foreach (var item in c.Groups[1].Value.Split(','))
                    {
                        var it = item.Trim();
                        var r = Regex.Match(it, @"^(-?\d+)\s+To\s+(-?\d+)$", RegexOptions.IgnoreCase);
                        if (Regex.IsMatch(it, @"^-?\d+$")) used.TryAdd(int.Parse(it), lineNo);
                        else if (r.Success && int.Parse(r.Groups[2].Value) - int.Parse(r.Groups[1].Value) <= 10000)
                            for (int n = int.Parse(r.Groups[1].Value); n <= int.Parse(r.Groups[2].Value); n++) used.TryAdd(n, lineNo);
                        else manual.Add($"satır {lineNo}: Case {it}");
                    }
                    continue;
                }
                if (Regex.IsMatch(s, @"^(If|ElseIf)\b", RegexOptions.IgnoreCase))
                {
                    var cond = Regex.Replace(s, @"\bThen\b.*$", "", RegexOptions.IgnoreCase);
                    // (?<![.\w]) skips obj.Index and ListIndex
                    foreach (Match e in Regex.Matches(cond, @"(?<![.\w])Index\s*=\s*(-?\d+)\b", RegexOptions.IgnoreCase))
                        used.TryAdd(int.Parse(e.Groups[1].Value), lineNo);
                    if (Regex.IsMatch(cond, @"(?<![.\w])Index\s*=\s*(?!-?\d+\b)\S", RegexOptions.IgnoreCase)
                        || Regex.IsMatch(cond, @"(?<![.\w])Index\s*(<>|<|>|<=|>=)", RegexOptions.IgnoreCase))
                        manual.Add($"satır {lineNo}: {cond.Trim()}");
                }
            }
        }
    done:
        if (used.Count == 0 && manual.Count == 0) return null;   // Index used generically, e.g. mnuX(Index).Caption

        var sb = new StringBuilder();
        var missing = used.Keys.Where(n => !defined.Contains(n)).ToList();
        if (missing.Count > 0)
            sb.Append($"\r\n    Kodda kullanılıp formda OLMAYAN index'ler ({missing.Count}): {Ranges(missing, used)}");
        if (used.Count > 0)
        {
            var unhandled = defined.Where(n => !used.ContainsKey(n)).ToList();
            if (unhandled.Count > 0)
                sb.Append($"\r\n    Formda olup kodda karşılığı OLMAYAN index'ler ({unhandled.Count}): {Ranges(unhandled)}"
                    + (caseElse ? "  (Case Else'e düşer)" : ""));
        }
        foreach (var mline in manual) sb.Append($"\r\n    Elle kontrol: {mline}");
        return sb.Length == 0 ? null : sb.ToString();
    }

    // "0-9, 12, 15-20"; with lines: "10-49 (s.55)"
    static string Ranges(IEnumerable<int> values, IDictionary<int, int>? lines = null)
    {
        var parts = new List<string>();
        var list = values.ToList();
        for (int i = 0; i < list.Count; )
        {
            int j = i;
            while (j + 1 < list.Count && list[j + 1] == list[j] + 1) j++;
            var p = i == j ? $"{list[i]}" : $"{list[i]}-{list[j]}";
            if (lines != null) p += $" (s.{lines[list[i]]})";
            parts.Add(p);
            i = j + 1;
        }
        return string.Join(", ", parts);
    }

    static string StripComment(string line)
    {
        bool inStr = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inStr = !inStr;
            else if (line[i] == '\'' && !inStr) return line[..i];
        }
        return Regex.IsMatch(line, @"^\s*Rem\b", RegexOptions.IgnoreCase) ? "" : line;
    }

    static IEnumerable<string> SplitStatements(string line)
    {
        bool inStr = false;
        int start = 0;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inStr = !inStr;
            else if (line[i] == ':' && !inStr) { yield return line[start..i]; start = i + 1; }
        }
        yield return line[start..];
    }

    static void ScanFile(string file, string rel, Encoding enc, List<string> orphans, List<string> indexIssues,
        List<string> indexUsage, List<string> possible, ref int handlerCount)
    {
        var lines = File.ReadAllLines(file, enc);

        // designer: control name -> Index values (empty = not a control array)
        var controls = new Dictionary<string, SortedSet<int>>(StringComparer.OrdinalIgnoreCase);
        var stack = new Stack<string>();
        int propDepth = 0, codeStart = 0;
        bool started = false;
        for (int i = 0; i < lines.Length; i++)
        {
            var t = lines[i].Trim();
            if (t.StartsWith("BeginProperty", StringComparison.OrdinalIgnoreCase)) { propDepth++; continue; }
            if (t.StartsWith("EndProperty", StringComparison.OrdinalIgnoreCase)) { propDepth--; continue; }
            if (propDepth > 0) continue;
            var b = Regex.Match(t, @"^Begin\s+(\S+)\s+(\w+)", RegexOptions.IgnoreCase);
            if (b.Success)
            {
                var name = b.Groups[2].Value;
                if (started && !controls.ContainsKey(name)) controls[name] = [];
                stack.Push(name);
                started = true;
                continue;
            }
            if (started && t.Equals("End", StringComparison.OrdinalIgnoreCase))
            {
                stack.Pop();
                if (stack.Count == 0) { codeStart = i + 1; break; }
                continue;
            }
            var ix = Regex.Match(t, @"^Index\s*=\s*(-?\d+)", RegexOptions.IgnoreCase);
            if (stack.Count > 1 && ix.Success)
                controls[stack.Peek()].Add(int.Parse(ix.Groups[1].Value));
        }

        // code: join line continuations, keep the first physical line number
        var logical = new List<(int Line, string Text)>();
        for (int i = codeStart; i < lines.Length; i++)
        {
            int start = i;
            var text = lines[i];
            while (Regex.IsMatch(text, @"\s_\s*$") && i + 1 < lines.Length)
                text = Regex.Replace(text, @"\s_\s*$", " ") + lines[++i].TrimStart();
            logical.Add((start + 1, text));
        }

        var withEvents = new List<string>();
        var implements = new List<string>();
        foreach (var (_, text) in logical)
        {
            foreach (Match w in Regex.Matches(text, @"\bWithEvents\s+(\w+)", RegexOptions.IgnoreCase)) withEvents.Add(w.Groups[1].Value);
            var im = Regex.Match(text, @"^\s*Implements\s+(\w+)", RegexOptions.IgnoreCase);
            if (im.Success) implements.Add(im.Groups[1].Value);
        }

        for (int li = 0; li < logical.Count; li++)
        {
            var (lineNo, text) = logical[li];
            var m = ProcRx.Match(text);
            if (!m.Success) continue;
            var proc = m.Groups[2].Value;
            int us = proc.LastIndexOf('_');
            if (us <= 0 || us == proc.Length - 1) continue;
            var obj = proc[..us];
            var evt = proc[(us + 1)..];
            if (SelfPrefixes.Contains(obj, StringComparer.OrdinalIgnoreCase)) continue;
            if (implements.Any(x => proc.StartsWith(x + "_", StringComparison.OrdinalIgnoreCase))) continue;
            if (withEvents.Any(x => obj.Equals(x, StringComparison.OrdinalIgnoreCase))) continue;

            bool hasIndex = Regex.IsMatch(m.Groups[3].Value, @"^\s*(ByVal\s+|ByRef\s+)?Index\s+As\s+Integer\b", RegexOptions.IgnoreCase);
            string where = $"{rel}:{lineNo}";

            if (controls.TryGetValue(obj, out var indices))
            {
                handlerCount++;
                bool isArray = indices.Count > 0;
                if (isArray && hasIndex)
                {
                    var usage = AnalyzeIndexUsage(logical, li, indices);
                    if (usage != null) indexUsage.Add($"{where}  {proc}  (formdaki index'ler: {Ranges(indices)})" + usage);
                }
                else if (isArray && !hasIndex)
                    indexIssues.Add($"{where}  {proc}  -> {obj} kontrol dizisi (Index'li), ama prosedürde 'Index As Integer' yok");
                else if (!isArray && hasIndex)
                    indexIssues.Add($"{where}  {proc}  -> {obj} tekil kontrol (Index yok), ama prosedür 'Index As Integer' alıyor");
                continue;
            }

            var arr = hasIndex ? " (Index'li)" : "";
            if (KnownEvents.Contains(evt))
            {
                handlerCount++;
                orphans.Add($"{where}  {proc}{arr}  -> formda '{obj}' adlı kontrol yok");
            }
            else
                possible.Add($"{where}  {proc}{arr}  -> '{obj}' kontrol değil, '{evt}' bilinen olay adı değil");
        }
    }
}
