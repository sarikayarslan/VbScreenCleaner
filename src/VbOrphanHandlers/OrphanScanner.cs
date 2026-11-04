// Finds event-handler procedures (xxx_Click, xxx_Change, ...) in VB5/VB6 forms whose
// control does not exist in the form designer, plus control-array Index mismatches.
// Read-only: VB files are never written.
using System.Text;
using System.Text.RegularExpressions;

namespace VbOrphanHandlers;

public static class OrphanScanner
{
    static readonly string[] FormExts = [".frm", ".ctl", ".pag", ".dob"];
    static readonly string[] CodeExts = [".frm", ".ctl", ".pag", ".dob", ".bas", ".cls"];
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
        var codeFiles = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);   // forms + .bas/.cls, searched for calls
        foreach (var input in inputs)
        {
            if (Directory.Exists(input))
            {
                foreach (var f in Directory.EnumerateFiles(input, "*.*", SearchOption.AllDirectories))
                {
                    if (FormExts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)) files.Add(Path.GetFullPath(f));
                    if (CodeExts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)) codeFiles.Add(Path.GetFullPath(f));
                }
            }
            else if (File.Exists(input) && input.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase))
            {
                var baseDir = Path.GetDirectoryName(Path.GetFullPath(input))!;
                foreach (var line in File.ReadAllLines(input, enc))
                {
                    var m = Regex.Match(line, @"^(Form|UserControl|PropertyPage|UserDocument|Module|Class)=(.+)$", RegexOptions.IgnoreCase);
                    if (!m.Success) continue;
                    var rel = m.Groups[2].Value.Trim();
                    if (rel.Contains(';')) rel = rel[(rel.IndexOf(';') + 1)..].Trim();
                    var full = Path.GetFullPath(Path.Combine(baseDir, rel));
                    if (!File.Exists(full)) { warnings.Add($".vbp içindeki dosya bulunamadı: {full}"); continue; }
                    codeFiles.Add(full);
                    if (FormExts.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase)) files.Add(full);
                }
            }
            else if (File.Exists(input) && CodeExts.Contains(Path.GetExtension(input), StringComparer.OrdinalIgnoreCase))
            {
                codeFiles.Add(Path.GetFullPath(input));
                if (FormExts.Contains(Path.GetExtension(input), StringComparer.OrdinalIgnoreCase)) files.Add(Path.GetFullPath(input));
            }
            else throw new FileNotFoundException($"Bulunamadı: {input}");
        }

        var root = inputs.Select(p => Directory.Exists(p) ? Path.GetFullPath(p) : Path.GetDirectoryName(Path.GetFullPath(p))!).First();
        var orphans = new List<string>();
        var possible = new List<string>();
        var indexIssues = new List<string>();
        var indexUsage = new List<string>();
        var constAccess = new List<string>();
        int handlerCount = 0;

        var calls = BuildCallIndex(codeFiles, root, enc);
        foreach (var file in files)
            ScanFile(file, Path.GetRelativePath(root, file), enc, calls, orphans, indexIssues, indexUsage, constAccess, possible, ref handlerCount);

        var report = new StringBuilder();
        report.AppendLine("VB5 Sahipsiz Olay Prosedürleri Raporu");
        report.AppendLine($"Kaynak: {string.Join("; ", inputs)}");
        report.AppendLine($"Taranan form dosyası: {files.Count}, çağrı aranan dosya (form/modül/sınıf): {codeFiles.Count}, olay prosedürü: {handlerCount}");
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
        Section("3. Kontrol dizilerinde index kullanımı (Select Case Index, If Index = n, kontrol(Index ± n))", indexUsage);
        Section("4. Formda olmayan sabit index erişimleri (ör. mnuX(15))", constAccess);
        Section("5. Kontrol olmayan nesne_Ad prosedürleri (bilinmeyen olay adı, elle kontrol edin)", possible);
        report.AppendLine("Not: Form_/MDIForm_/UserControl_ olayları, WithEvents değişkenleri ve Implements arayüz metotları raporlanmaz.");
        return report.ToString();
    }

    const string IndexOffset = @"(?<![.\w])Index\s*(?:([+-])\s*(\d+))?";
    static readonly Regex IndexAccessRx = new(@"(?<![.\w])(\w+)\s*\(\s*" + IndexOffset + @"\s*\)", RegexOptions.IgnoreCase);
    static readonly Regex AnyAccessWithIndexRx = new(@"(?<![.\w])(\w+)\s*\(([^()]*(?<![.\w])Index\b[^()]*)\)", RegexOptions.IgnoreCase);
    static readonly Regex AnyCallRx = new(@"\w+\s*\([^()]*\)");
    static readonly Regex BareIndexRx = new(@"(?<![.\w])Index\b", RegexOptions.IgnoreCase);

    static int Offset(Match m, int signGroup) =>
        m.Groups[signGroup].Success ? (m.Groups[signGroup].Value == "-" ? -1 : 1) * int.Parse(m.Groups[signGroup + 1].Value) : 0;

    static string WithOffset(int off) => off == 0 ? "Index" : off > 0 ? $"Index + {off}" : $"Index - {-off}";

    // Compares Index values used in the handler body (Select Case Index [± n] / If Index [± n] = m) with the
    // designer's Index values, and checks array accesses like ctl(Index ± n). Returns report lines or null.
    static string? AnalyzeIndexUsage(List<(int Line, string Text)> logical, int headerAt, string array,
        Dictionary<string, SortedSet<int>> controls, HashSet<string> dynamicArrays, List<CallSite> sure, List<CallSite> maybe)
    {
        var defined = controls[array];
        // Index values the handler is called with directly (mnuX_Click 15)
        var calledWith = new SortedDictionary<int, List<CallSite>>();
        var otherCalls = new List<CallSite>(maybe);
        foreach (var c in sure)
            if (int.TryParse(c.FirstArg, out var n)) { if (!calledWith.TryGetValue(n, out var l)) calledWith[n] = l = []; l.Add(c); }
            else otherCalls.Add(c);
        var domain = new SortedSet<int>(defined.Concat(calledWith.Keys));   // every Index value the body can see
        var used = new SortedDictionary<int, int>();   // index -> first line
        var manual = new List<string>();
        var risks = new List<string>();
        var selectStack = new Stack<int?>();           // Index offset of Select Case Index ± n; null = other select
        bool caseElse = false;
        for (int i = headerAt + 1; i < logical.Count; i++)
        {
            var (lineNo, raw) = logical[i];
            foreach (var stmt in SplitStatements(StripComment(raw)))
            {
                var s = StripStrings(stmt).Trim();
                if (Regex.IsMatch(s, @"^End\s+Sub\b", RegexOptions.IgnoreCase)) goto done;

                // array accesses: ctl(Index ± n)
                foreach (Match a in IndexAccessRx.Matches(s))
                {
                    var target = a.Groups[1].Value;
                    if (!controls.TryGetValue(target, out var targetIdx) || targetIdx.Count == 0) continue;
                    int off = Offset(a, 2);
                    var bad = domain.Where(n => !targetIdx.Contains(n + off)).ToList();
                    if (bad.Count == 0) continue;
                    var dyn = dynamicArrays.Contains(target) ? "  (Load ile çalışma zamanında eleman ekleniyor, yüklenmiş olabilir)" : "";
                    risks.Add($"satır {lineNo}: {target}({WithOffset(off)}) -> formda olmayan: {target}({Ranges(bad.Select(n => n + off))}), "
                        + $"index {Ranges(bad)} için{dyn}");
                }
                foreach (Match a in AnyAccessWithIndexRx.Matches(s))
                    if (controls.TryGetValue(a.Groups[1].Value, out var t) && t.Count > 0 && !IndexAccessRx.IsMatch(a.Value))
                        manual.Add($"satır {lineNo}: {a.Value}");

                var sel = Regex.Match(s, @"^Select\s+Case\s+(.+)$", RegexOptions.IgnoreCase);
                if (sel.Success)
                {
                    var expr = sel.Groups[1].Value.Trim();
                    var plain = Regex.Match(expr, "^" + IndexOffset + @"$", RegexOptions.IgnoreCase);
                    if (plain.Success) selectStack.Push(Offset(plain, 1));
                    else
                    {
                        if (BareIndexRx.IsMatch(AnyCallRx.Replace(expr, ""))) manual.Add($"satır {lineNo}: {s}");
                        selectStack.Push(null);
                    }
                    continue;
                }
                if (Regex.IsMatch(s, @"^End\s+Select\b", RegexOptions.IgnoreCase)) { if (selectStack.Count > 0) selectStack.Pop(); continue; }
                var c = Regex.Match(s, @"^Case\s+(.+)$", RegexOptions.IgnoreCase);
                if (c.Success && selectStack.Count > 0 && selectStack.Peek() is int selOff)
                {
                    if (Regex.IsMatch(c.Groups[1].Value, @"^Else\b", RegexOptions.IgnoreCase)) { caseElse = true; continue; }
                    foreach (var item in c.Groups[1].Value.Split(','))
                    {
                        var it = item.Trim();
                        var r = Regex.Match(it, @"^(-?\d+)\s+To\s+(-?\d+)$", RegexOptions.IgnoreCase);
                        if (Regex.IsMatch(it, @"^-?\d+$")) used.TryAdd(int.Parse(it) - selOff, lineNo);
                        else if (r.Success && int.Parse(r.Groups[2].Value) - int.Parse(r.Groups[1].Value) <= 10000)
                            for (int n = int.Parse(r.Groups[1].Value); n <= int.Parse(r.Groups[2].Value); n++) used.TryAdd(n - selOff, lineNo);
                        else manual.Add($"satır {lineNo}: Case {it}" + (selOff != 0 ? $"  (Select Case {WithOffset(selOff)})" : ""));
                    }
                    continue;
                }
                if (Regex.IsMatch(s, @"^(If|ElseIf)\b", RegexOptions.IgnoreCase))
                {
                    var cond = Regex.Replace(s, @"\bThen\b.*$", "", RegexOptions.IgnoreCase);
                    // Index [± n] = m  -> index m ∓ n; (?<![.\w]) skips obj.Index and ListIndex
                    var cmp = new Regex(IndexOffset + @"\s*=\s*(-?\d+)\b(?!\s*[-+*/\^.])", RegexOptions.IgnoreCase);
                    foreach (Match e in cmp.Matches(cond))
                        used.TryAdd(int.Parse(e.Groups[3].Value) - Offset(e, 1), lineNo);
                    var rest = AnyCallRx.Replace(cmp.Replace(cond, ""), "");
                    if (BareIndexRx.IsMatch(rest)) manual.Add($"satır {lineNo}: {cond.Trim()}");
                }
            }
        }
    done:
        // nothing compared: Index used generically, e.g. mnuX(Index).Caption
        if (used.Count == 0 && manual.Count == 0 && risks.Count == 0 && otherCalls.Count == 0
            && calledWith.Keys.All(defined.Contains)) return null;

        var sb = new StringBuilder();
        var missing = used.Keys.Where(n => !defined.Contains(n) && !calledWith.ContainsKey(n)).ToList();
        foreach (var (n, cs) in calledWith.Where(kv => !defined.Contains(kv.Key)))
            sb.Append($"\r\n    Formda yok ama ÇAĞRIYLA kullanılıyor, ölü değil: index {n} ({Sites(cs)})"
                + (used.Count > 0 && !used.ContainsKey(n) ? (caseElse ? "  (Case Else'e düşer)" : "  (kodda karşılığı yok)") : ""));
        if (missing.Count > 0)
            sb.Append($"\r\n    Kodda kullanılıp formda OLMAYAN index'ler ({missing.Count}): {Ranges(missing, used)}"
                + (dynamicArrays.Contains(array) ? "  (Load ile çalışma zamanında eleman ekleniyor, yüklenmiş olabilir)" : ""));
        if (used.Count > 0)
        {
            var unhandled = defined.Where(n => !used.ContainsKey(n)).ToList();
            if (unhandled.Count > 0)
                sb.Append($"\r\n    Formda olup kodda karşılığı OLMAYAN index'ler ({unhandled.Count}): {Ranges(unhandled)}"
                    + (caseElse ? "  (Case Else'e düşer)" : ""));
        }
        foreach (var r in risks) sb.Append($"\r\n    Çalışma zamanı hatası riski: {r}");
        foreach (var mline in manual.Distinct()) sb.Append($"\r\n    Elle kontrol: {mline}");
        foreach (var c in otherCalls)
            sb.Append($"\r\n    Elle kontrol: {c.Where}: " + (c.ByName ? "CallByName ile çağrılıyor (olası)"
                : maybe.Contains(c) ? $"olası çağrı, '{c.Qualifier}.' üzerinden (form değişkeni olabilir), index: {(c.FirstArg.Length == 0 ? "yok" : c.FirstArg)}"
                : c.FirstArg.Length == 0 ? "index'siz çağrılıyor" : $"değişken/ifade ile çağrılıyor ({c.FirstArg})"));
        return sb.Length == 0 ? null : sb.ToString();
    }

    // A direct call of a procedure: mnuX_Click 15 / Call mnuX_Click(15) / frmAna.mnuX_Click 15 / CallByName(f, "mnuX_Click", ...)
    record CallSite(string File, string Rel, int Line, string? Qualifier, string FirstArg, bool ByName)
    {
        public string Where => $"{Rel}:{Line}";
    }

    static readonly Regex CallStmtRx = new(@"^\s*(?:Call\s+)?(?:(\w+)\s*\.\s*)?([A-Za-z]\w*_\w+)\b(.*)$", RegexOptions.IgnoreCase);
    static readonly Regex DeclRx = new(@"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property|Declare|Event|End)\b", RegexOptions.IgnoreCase);

    static Dictionary<string, List<CallSite>> BuildCallIndex(IEnumerable<string> codeFiles, string root, Encoding enc)
    {
        var calls = new Dictionary<string, List<CallSite>>(StringComparer.OrdinalIgnoreCase);
        void Add(string proc, CallSite c)
        {
            if (!calls.TryGetValue(proc, out var list)) calls[proc] = list = [];
            list.Add(c);
        }
        foreach (var file in codeFiles)
        {
            var lines = File.ReadAllLines(file, enc);
            int start = Array.FindIndex(lines, l => l.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase));
            var rel = Path.GetRelativePath(root, file);
            foreach (var (lineNo, text) in Logical(lines, Math.Max(start, 0)))
            {
                var raw = StripComment(text);
                foreach (Match b in Regex.Matches(raw, @"\bCallByName\s*\(\s*([^,]+?)\s*,\s*""(\w+)""", RegexOptions.IgnoreCase))
                    Add(b.Groups[2].Value, new CallSite(file, rel, lineNo, b.Groups[1].Value, "?", true));
                var code = StripStrings(raw);
                if (DeclRx.IsMatch(code)) continue;
                foreach (var stmt in SplitStatements(code))
                {
                    // single-line If ... Then call / Else call
                    var st = Regex.Replace(stmt, @"^\s*(?:(?:ElseIf|If)\b.*?\bThen\b|Else\b)", "", RegexOptions.IgnoreCase);
                    var m = CallStmtRx.Match(st);
                    if (!m.Success) continue;
                    var rest = m.Groups[3].Value;
                    if (Regex.IsMatch(rest, @"^\s*(=|\.|!)")) continue;   // assignment / member access, not a call
                    Add(m.Groups[2].Value, new CallSite(file, rel, lineNo, m.Groups[1].Success ? m.Groups[1].Value : null, FirstArg(rest), false));
                }
            }
        }
        return calls;
    }

    // first argument of "(a, b)" or " a, b"; "" when there is none
    static string FirstArg(string rest)
    {
        var r = rest.Trim();
        if (r.StartsWith('(')) r = r[1..];
        int depth = 0;
        for (int i = 0; i < r.Length; i++)
        {
            if (r[i] == '(') depth++;
            else if (r[i] == ')' && depth-- == 0) return r[..i].Trim();
            else if (r[i] == ',' && depth == 0) return r[..i].Trim();
        }
        return r.Trim();
    }

    // joins line continuations, keeps the first physical line number
    static List<(int Line, string Text)> Logical(string[] lines, int start)
    {
        var logical = new List<(int Line, string Text)>();
        for (int i = start; i < lines.Length; i++)
        {
            int first = i;
            var text = lines[i];
            while (Regex.IsMatch(text, @"\s_\s*$") && i + 1 < lines.Length)
                text = Regex.Replace(text, @"\s_\s*$", " ") + lines[++i].TrimStart();
            logical.Add((first + 1, text));
        }
        return logical;
    }

    static string Sites(IEnumerable<CallSite> sites)
    {
        var list = sites.Select(c => c.Where).Distinct().ToList();
        return string.Join(", ", list.Take(5)) + (list.Count > 5 ? $" ... (+{list.Count - 5})" : "");
    }

    static string StripStrings(string s) => Regex.Replace(s, "\"[^\"]*\"", "\"\"");

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

    static void ScanFile(string file, string rel, Encoding enc, Dictionary<string, List<CallSite>> calls, List<string> orphans, List<string> indexIssues,
        List<string> indexUsage, List<string> constAccess, List<string> possible, ref int handlerCount)
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

        var logical = Logical(lines, codeStart);
        var vbName = lines.Select(l => Regex.Match(l, @"^Attribute\s+VB_Name\s*=\s*""(\w+)""", RegexOptions.IgnoreCase))
            .FirstOrDefault(x => x.Success)?.Groups[1].Value ?? Path.GetFileNameWithoutExtension(file);

        // direct calls of a procedure of this form: same file (unqualified / Me. / Form.), other files only Form.proc
        // when the procedure is not Private. Other qualifiers (form instance variables) and CallByName are "possible".
        (List<CallSite> Sure, List<CallSite> Possible) CallsOf(string proc, bool isPrivate)
        {
            var sure = new List<CallSite>();
            var maybe = new List<CallSite>();
            if (!calls.TryGetValue(proc, out var sites)) return (sure, maybe);
            foreach (var c in sites)
            {
                bool same = string.Equals(c.File, file, StringComparison.OrdinalIgnoreCase);
                bool selfQ = c.Qualifier == null || c.Qualifier.Equals("Me", StringComparison.OrdinalIgnoreCase)
                    || c.Qualifier.Equals(vbName, StringComparison.OrdinalIgnoreCase);
                if (c.ByName) { if (same || !isPrivate) maybe.Add(c); }
                else if (same && selfQ) sure.Add(c);
                else if (!same && !isPrivate && c.Qualifier != null)
                    (c.Qualifier.Equals(vbName, StringComparison.OrdinalIgnoreCase) ? sure : maybe).Add(c);
            }
            return (sure, maybe);
        }

        var withEvents = new List<string>();
        var implements = new List<string>();
        foreach (var (_, text) in logical)
        {
            foreach (Match w in Regex.Matches(text, @"\bWithEvents\s+(\w+)", RegexOptions.IgnoreCase)) withEvents.Add(w.Groups[1].Value);
            var im = Regex.Match(text, @"^\s*Implements\s+(\w+)", RegexOptions.IgnoreCase);
            if (im.Success) implements.Add(im.Groups[1].Value);
        }

        // arrays that get elements at runtime (Load mnuX(i)), and constant accesses like mnuX(15)
        var dynamicArrays = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var code = logical.Select(l => (l.Line, Text: StripStrings(StripComment(l.Text)))).ToList();
        foreach (var (_, text) in code)
            foreach (Match ld in Regex.Matches(text, @"\bLoad\s+(\w+)\s*\(", RegexOptions.IgnoreCase)) dynamicArrays.Add(ld.Groups[1].Value);
        foreach (var (lineNo, text) in code)
            foreach (Match a in Regex.Matches(text, @"(?<![.\w])(\w+)\s*\(\s*(-?\d+)\s*\)", RegexOptions.IgnoreCase))
            {
                var name = a.Groups[1].Value;
                if (!controls.TryGetValue(name, out var idx) || idx.Count == 0 || idx.Contains(int.Parse(a.Groups[2].Value))) continue;
                if (Regex.IsMatch(text[..a.Index], @"\bLoad\s+$", RegexOptions.IgnoreCase)) continue;
                constAccess.Add($"{rel}:{lineNo}  {name}({a.Groups[2].Value})  -> formda yok (formdaki index'ler: {Ranges(idx)})"
                    + (dynamicArrays.Contains(name) ? "  (Load ile çalışma zamanında eleman ekleniyor, yüklenmiş olabilir)" : ""));
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

            bool isPrivate = m.Groups[1].Value.Equals("Private", StringComparison.OrdinalIgnoreCase);
            var (sure, maybe) = CallsOf(proc, isPrivate);
            bool hasIndex = Regex.IsMatch(m.Groups[3].Value, @"^\s*(ByVal\s+|ByRef\s+)?Index\s+As\s+Integer\b", RegexOptions.IgnoreCase);
            string where = $"{rel}:{lineNo}";

            if (controls.TryGetValue(obj, out var indices))
            {
                handlerCount++;
                bool isArray = indices.Count > 0;
                if (isArray && hasIndex)
                {
                    var usage = AnalyzeIndexUsage(logical, li, obj, controls, dynamicArrays, sure, maybe);
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
                var called = sure.Count > 0 ? $"  [ÇAĞRILIYOR, ölü değil: {Sites(sure)}]" : "";
                if (maybe.Count > 0) called += $"  [olası çağrı, elle kontrol: {Sites(maybe)}]";
                orphans.Add($"{where}  {proc}{arr}  -> formda '{obj}' adlı kontrol yok{called}");
            }
            else
                possible.Add($"{where}  {proc}{arr}  -> '{obj}' kontrol değil, '{evt}' bilinen olay adı değil");
        }
    }
}
