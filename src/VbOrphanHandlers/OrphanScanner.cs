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
        int handlerCount = 0;

        foreach (var file in files)
            ScanFile(file, Path.GetRelativePath(root, file), enc, orphans, indexIssues, possible, ref handlerCount);

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
        Section("3. Kontrol olmayan nesne_Ad prosedürleri (bilinmeyen olay adı, elle kontrol edin)", possible);
        report.AppendLine("Not: Form_/MDIForm_/UserControl_ olayları, WithEvents değişkenleri ve Implements arayüz metotları raporlanmaz.");
        return report.ToString();
    }

    static void ScanFile(string file, string rel, Encoding enc, List<string> orphans, List<string> indexIssues,
        List<string> possible, ref int handlerCount)
    {
        var lines = File.ReadAllLines(file, enc);

        // designer: control name -> is it a control array?
        var controls = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
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
                if (started && !controls.ContainsKey(name)) controls[name] = false;
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
            if (stack.Count > 1 && Regex.IsMatch(t, @"^Index\s*=\s*-?\d+", RegexOptions.IgnoreCase))
                controls[stack.Peek()] = true;
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

        foreach (var (lineNo, text) in logical)
        {
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

            if (controls.TryGetValue(obj, out var isArray))
            {
                handlerCount++;
                if (isArray && !hasIndex)
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
