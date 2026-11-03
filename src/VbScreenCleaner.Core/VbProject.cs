using System.Text;
using System.Text.RegularExpressions;

namespace VbScreenCleaner.Core;

public sealed class VbProject
{
    // VB5 files are ANSI (e.g. Windows-1254). Latin1 maps every byte to one char, so a
    // read/write round trip keeps the original bytes untouched.
    public static readonly Encoding FileEncoding = Encoding.Latin1;

    static readonly Dictionary<string, VbFileKind> KindByExtension = new(StringComparer.OrdinalIgnoreCase)
    {
        [".frm"] = VbFileKind.Form,
        [".ctl"] = VbFileKind.UserControl,
        [".pag"] = VbFileKind.PropertyPage,
        [".dob"] = VbFileKind.UserDocument,
        [".dsr"] = VbFileKind.Designer,
        [".bas"] = VbFileKind.Module,
        [".cls"] = VbFileKind.Class,
    };

    public required string RootDir { get; init; }
    public List<string> ProjectFiles { get; } = new();
    public List<VbSourceFile> Files { get; } = new();
    public List<string> Warnings { get; } = new();

    public static VbProject Load(string rootDir)
    {
        var root = Path.GetFullPath(rootDir).TrimEnd(Path.DirectorySeparatorChar);
        if (!Directory.Exists(root)) throw new DirectoryNotFoundException(root);

        var project = new VbProject { RootDir = root };
        var paths = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var vbp in Directory.EnumerateFiles(root, "*.vbp", SearchOption.AllDirectories))
        {
            project.ProjectFiles.Add(vbp);
            foreach (var p in ReadVbp(vbp, project.Warnings))
                if (seen.Add(p)) paths.Add(p);
        }

        if (project.ProjectFiles.Count == 0)
        {
            project.Warnings.Add("Hiç .vbp bulunamadı; klasördeki tüm VB dosyaları işleniyor.");
            foreach (var p in Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories))
                if (KindByExtension.ContainsKey(Path.GetExtension(p)) && seen.Add(p)) paths.Add(p);
        }

        foreach (var path in paths)
        {
            if (!File.Exists(path))
            {
                project.Warnings.Add($"Projede tanımlı dosya bulunamadı: {path}");
                continue;
            }
            var file = LoadFile(path, root);
            if (!file.IsInsideRoot)
                project.Warnings.Add($"Dosya proje klasörünün dışında, yalnızca okunacak ve hedefe kopyalanmayacak: {path}");
            project.Files.Add(file);
        }

        foreach (var f in project.Files) VbParser.Parse(f);
        return project;
    }

    static IEnumerable<string> ReadVbp(string vbpPath, List<string> warnings)
    {
        var dir = Path.GetDirectoryName(vbpPath)!;
        foreach (var raw in File.ReadAllLines(vbpPath, FileEncoding))
        {
            int eq = raw.IndexOf('=');
            if (eq <= 0) continue;
            var key = raw[..eq].Trim();
            var value = raw[(eq + 1)..].Trim();
            string? rel = key.ToLowerInvariant() switch
            {
                "form" or "usercontrol" or "propertypage" or "userdocument" or "designer" => value,
                // Module=Name; path   Class=Name; path
                "module" or "class" => value.Contains(';') ? value[(value.IndexOf(';') + 1)..].Trim() : value,
                _ => null,
            };
            if (string.IsNullOrEmpty(rel)) continue;
            rel = rel.Trim('"');
            if (!KindByExtension.ContainsKey(Path.GetExtension(rel)))
            {
                warnings.Add($"{Path.GetFileName(vbpPath)}: tanınmayan dosya türü atlandı: {rel}");
                continue;
            }
            yield return Path.GetFullPath(Path.Combine(dir, rel));
        }
    }

    public static VbSourceFile LoadFile(string path, string root)
    {
        var text = File.ReadAllText(path, FileEncoding);
        string newLine = text.Contains("\r\n") ? "\r\n" : "\n";
        bool endsWithNewLine = text.EndsWith('\n');
        var lines = text.Split('\n').Select(l => l.TrimEnd('\r')).ToList();
        if (endsWithNewLine) lines.RemoveAt(lines.Count - 1);

        var full = Path.GetFullPath(path);
        string? relative = full.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            ? full[(root.Length + 1)..]
            : null;

        return new VbSourceFile
        {
            FullPath = full,
            RelativePath = relative,
            Kind = KindByExtension[Path.GetExtension(path)],
            Lines = lines,
            NewLine = newLine,
            EndsWithNewLine = endsWithNewLine,
        };
    }
}

public static class VbParser
{
    static readonly Regex BeginRx = new(@"^\s*Begin\s+(\S+)\s+(\S+)\s*$", RegexOptions.Compiled);
    // Property names may be indexed/dotted, e.g. SSTab's  Tab(0).Control(1)=   "txtX"
    static readonly Regex PropRx = new(@"^\s*([A-Za-z_][\w.()]*)\s*=\s*(.*)$", RegexOptions.Compiled);
    static readonly Regex VbNameRx = new(@"^\s*Attribute\s+VB_Name\s*=\s*""([^""]*)""", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    static readonly HashSet<string> ProcModifiers = new(StringComparer.OrdinalIgnoreCase) { "Public", "Private", "Friend", "Static", "Global" };
    static readonly HashSet<string> NonLabelKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "Else", "End", "Loop", "Wend", "Next", "Do", "Return", "Stop", "Resume", "Beep", "Randomize", "Exit", "Cls", "Print",
    };

    public static void Parse(VbSourceFile f)
    {
        if (f.Kind is VbFileKind.Module or VbFileKind.Class) f.CodeStart = FindClassHeaderEnd(f);
        else ParseDesigner(f);

        f.ObjectName = f.Lines.Skip(f.Kind is VbFileKind.Module or VbFileKind.Class ? 0 : f.CodeStart)
            .Select(l => VbNameRx.Match(l)).FirstOrDefault(m => m.Success)?.Groups[1].Value
            ?? f.Root?.Name ?? Path.GetFileNameWithoutExtension(f.FullPath);

        ParseCode(f);

        if (f.HasDesigner)
        {
            foreach (var proc in f.Procs.Where(p => p.Kind.Equals("Sub", StringComparison.OrdinalIgnoreCase)))
            {
                int us = proc.Name.LastIndexOf('_');
                if (us > 0 && f.Groups.TryGetValue(proc.Name[..us], out var g))
                {
                    g.Handlers.Add(proc);
                    proc.HandlerOwner = g;
                }
            }
        }
    }

    static int FindClassHeaderEnd(VbSourceFile f)
    {
        if (f.Lines.Count == 0 || !f.Lines[0].TrimStart().StartsWith("VERSION", StringComparison.OrdinalIgnoreCase)) return 0;
        for (int i = 1; i < f.Lines.Count && i < 50; i++)
            if (f.Lines[i].Trim() == "END") return i + 1;
        return 0;
    }

    static void ParseDesigner(VbSourceFile f)
    {
        var stack = new Stack<ControlNode>();
        int propDepth = 0;
        for (int i = 0; i < f.Lines.Count; i++)
        {
            var line = f.Lines[i];
            var t = line.Trim();
            if (propDepth == 0 && BeginRx.Match(line) is { Success: true } m)
            {
                var parent = stack.Count > 0 ? stack.Peek() : null;
                var node = new ControlNode { TypeName = m.Groups[1].Value, Name = m.Groups[2].Value, BeginLine = i, Parent = parent };
                if (parent == null) f.Root = node;
                else parent.Children.Add(node);
                stack.Push(node);
            }
            else if (stack.Count == 0) continue;
            else if (t.StartsWith("BeginProperty", StringComparison.Ordinal)) propDepth++;
            else if (t.StartsWith("EndProperty", StringComparison.Ordinal)) propDepth--;
            else if (t == "End" && propDepth == 0)
            {
                var node = stack.Pop();
                node.EndLine = i;
                if (stack.Count == 0)
                {
                    f.CodeStart = i + 1;
                    break;
                }
            }
            else if (PropRx.Match(line) is { Success: true } pm)
            {
                if (propDepth == 0)
                {
                    stack.Peek().Properties[pm.Groups[1].Value] = pm.Groups[2].Value;
                    stack.Peek().PropertyLines[pm.Groups[1].Value] = i;
                }
                else stack.Peek().NestedProperties.Add((pm.Groups[1].Value, pm.Groups[2].Value));
            }
        }

        if (f.Root == null) return;
        foreach (var node in f.Root.Descendants())
        {
            if (node.Properties.TryGetValue("Index", out var idx) && int.TryParse(PropertyValue.Strip(idx), out var iv))
                node.Index = iv;
            if (!f.Groups.TryGetValue(node.Name, out var g))
                f.Groups[node.Name] = g = new ControlGroup { File = f, Name = node.Name };
            g.Elements.Add(node);
            node.Group = g;
        }
    }

    static void ParseCode(VbSourceFile f)
    {
        ProcInfo? cur = null;
        int depth = 0;
        bool afterExit = false;
        int i = f.CodeStart;
        while (i < f.Lines.Count)
        {
            if (f.Lines[i].TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase)) { i++; continue; }

            int first = i;
            var sb = new System.Text.StringBuilder();
            string comment = "";
            while (i < f.Lines.Count)
            {
                var (code, cm) = VbLexer.SplitComment(f.Lines[i]);
                var trimmed = code.TrimEnd();
                i++;
                if (cm.Length == 0 && (trimmed == "_" || (trimmed.EndsWith('_') && trimmed.Length > 1 && char.IsWhiteSpace(trimmed[^2]))))
                {
                    sb.Append(trimmed[..^1]).Append(' ');
                    continue;
                }
                sb.Append(code);
                comment = cm;
                break;
            }

            var firstLine = f.Lines[first];
            var codeText = sb.ToString();
            var ll = new LogicalLine
            {
                FirstLine = first,
                LastLine = i - 1,
                Indent = firstLine[..(firstLine.Length - firstLine.TrimStart().Length)],
                Code = codeText,
                Comment = comment,
                Tokens = VbLexer.Tokenize(codeText),
            };
            f.LogicalLines.Add(ll);
            if (ll.Tokens.Count == 0) continue;

            var toks = ll.Tokens;
            var segments = VbLexer.SplitColons(toks, 0, toks.Count);
            int segIdx = 0;

            // Leading line label, kept verbatim: "Label:" / "10:" or a bare line number "10 stmt".
            if (segments.Count > 0 && segments[0] is var (s0, e0))
            {
                if (e0 - s0 == 1 && e0 < toks.Count && toks[e0].Kind == TokKind.Colon
                    && (toks[s0].Kind == TokKind.Number || (toks[s0].Kind == TokKind.Ident && !NonLabelKeywords.Contains(toks[s0].Text))))
                {
                    ll.Label = codeText[toks[s0].Start..toks[e0].End];
                    segIdx = 1;
                }
                else if (toks[s0].Kind == TokKind.Number)
                {
                    ll.Label = codeText[toks[s0].Start..toks[s0].End];
                    if (e0 - s0 == 1) segIdx = 1;
                    else segments[0] = (s0 + 1, e0);
                }
            }

            var units = new List<CodeUnit>();
            for (; segIdx < segments.Count; segIdx++)
            {
                var (s, e) = segments[segIdx];
                if (toks[s].Is("If"))
                {
                    int thenIdx = VbLexer.FindIdent(toks, s, e, "Then");
                    bool codeAfterThen = thenIdx >= 0
                        && (thenIdx < e - 1 || toks.Skip(e).Any(t => t.Kind != TokKind.Colon)); // "If x Then: y" too
                    if (codeAfterThen)
                    {
                        // Single-line If: everything to the end of the line belongs to it.
                        units.Add(new CodeUnit { File = f, Line = ll, TokStart = s, TokEnd = toks.Count, IsSingleLineIf = true });
                        break;
                    }
                }
                units.Add(new CodeUnit { File = f, Line = ll, TokStart = s, TokEnd = e });
            }

            foreach (var u in units)
            {
                ll.Units.Add(u);
                f.Units.Add(u);

                if (TryProcHeader(toks, u.TokStart, u.TokEnd, out var name, out var kind, out var isPrivate))
                {
                    cur = new ProcInfo { File = f, Name = name, Kind = kind, HeaderLine = ll.FirstLine, IsPrivate = isPrivate };
                    f.Procs.Add(cur);
                    cur.Header = u;
                    cur.Units.Add(u);
                    u.Proc = cur;
                    depth = 0;
                    afterExit = false;
                    continue;
                }

                u.Proc = cur;
                u.Depth = depth;
                u.AfterExit = afterExit;
                cur?.Units.Add(u);

                if (cur != null && IsEndProc(toks, u.TokStart, u.TokEnd))
                {
                    cur.End = u;
                    cur.EndLine = ll.LastLine;
                    cur = null;
                    depth = 0;
                    continue;
                }

                if (!u.IsSingleLineIf) depth = Math.Max(0, depth + DepthDelta(toks, u.TokStart, u.TokEnd));
                if (ContainsExit(toks, u.TokStart, u.TokEnd)) afterExit = true;
            }
        }
    }

    static bool TryProcHeader(List<Token> t, int s, int e, out string name, out string kind, out bool isPrivate)
    {
        name = kind = "";
        isPrivate = false;
        int p = s;
        while (p < e && t[p].Kind == TokKind.Ident && ProcModifiers.Contains(t[p].Text))
        {
            if (t[p].Is("Private")) isPrivate = true;
            p++;
        }
        if (p >= e || t[p].Kind != TokKind.Ident) return false;
        if (t[p].Is("Sub") || t[p].Is("Function"))
        {
            kind = t[p].Text;
            if (p + 1 < e && t[p + 1].Kind == TokKind.Ident) { name = t[p + 1].Text; return true; }
            return false;
        }
        if (t[p].Is("Property") && p + 2 < e && (t[p + 1].Is("Get") || t[p + 1].Is("Let") || t[p + 1].Is("Set")))
        {
            kind = "Property " + t[p + 1].Text;
            name = t[p + 2].Text;
            return true;
        }
        return false;
    }

    static bool IsEndProc(List<Token> t, int s, int e) =>
        e - s == 2 && t[s].Is("End") && (t[s + 1].Is("Sub") || t[s + 1].Is("Function") || t[s + 1].Is("Property"));

    static int DepthDelta(List<Token> t, int s, int e)
    {
        var k = t[s];
        if (k.Kind == TokKind.Other && k.Text == "#" && s + 1 < e)
        {
            if (t[s + 1].Is("If")) return 1;
            if (t[s + 1].Is("End")) return -1;
            return 0;
        }
        if (k.Kind != TokKind.Ident) return 0;
        if (k.Is("If") || k.Is("Select") || k.Is("For") || k.Is("Do") || k.Is("While") || k.Is("With")) return 1;
        if (k.Is("Next") || k.Is("Loop") || k.Is("Wend")) return -1;
        if (k.Is("End") && s + 1 < e && (t[s + 1].Is("If") || t[s + 1].Is("Select") || t[s + 1].Is("With"))) return -1;
        return 0;
    }

    static bool ContainsExit(List<Token> t, int s, int e)
    {
        if (e - s == 1 && t[s].Is("End")) return true;
        for (int i = s; i < e; i++)
        {
            if (t[i].Is("Exit") && i + 1 < e && (t[i + 1].Is("Sub") || t[i + 1].Is("Function") || t[i + 1].Is("Property"))) return true;
            if (t[i].Is("Return")) return true;
            if ((t[i].Is("GoTo") || t[i].Is("Resume")) && !(i >= 2 && t[i - 2].Is("On") && t[i - 1].Is("Error"))
                && !(i >= 1 && t[i - 1].Is("Error"))) return true;
        }
        return false;
    }
}
