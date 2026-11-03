namespace VbScreenCleaner.Core;

public sealed class CleanupOptions
{
    /// <summary>Also remove controls that are visible at design time but hidden unconditionally in Form_Load.</summary>
    public bool IncludeHiddenInFormLoad { get; set; } = true;

    /// <summary>Leave forms alone whose code touches the Controls collection (dynamic access by name/loop).</summary>
    public bool SkipFormsUsingControlsCollection { get; set; } = true;

    /// <summary>Group keys (<see cref="ControlGroup.Key"/>) the user chose to keep.</summary>
    public HashSet<string> ForcedKeep { get; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class FileResult
{
    public required VbSourceFile File { get; init; }
    public List<ControlNode> RemovalRoots { get; } = new();
    public List<ControlNode> KeptHidden { get; } = new();
    public List<ProcInfo> RemovedProcs { get; } = new();
    public List<CodeUnit> RemovedUnits { get; } = new();
    public List<string> Notes { get; } = new();

    public int RemovedControlCount => RemovalRoots.Sum(r => r.SelfAndDescendants().Count());
    public bool HasChanges => RemovalRoots.Count > 0 || RemovedProcs.Count > 0 || RemovedUnits.Count > 0;
}

public sealed class AnalysisResult
{
    public required VbProject Project { get; init; }
    public required CleanupOptions Options { get; init; }
    public List<FileResult> Files { get; } = new();
    public List<string> Warnings { get; } = new();
    public int Iterations { get; set; }

    public int TotalRemovedControls => Files.Sum(f => f.RemovedControlCount);
}

/// <summary>
/// Decides which controls can never become visible and are safe to delete:
/// hidden forever (design time or unconditionally in Form_Load, never shown by code)
/// and not referenced by any code that would survive the cleanup.
/// </summary>
public sealed class CleanupAnalyzer
{
    static readonly HashSet<string> SelfNames = new(StringComparer.OrdinalIgnoreCase) { "Me", "UserControl", "PropertyPage", "UserDocument" };
    static readonly HashSet<string> SafeMethods = new(StringComparer.OrdinalIgnoreCase) { "Move", "ZOrder", "Refresh", "Clear", "AddItem", "RemoveItem", "Cls" };
    static readonly HashSet<string> LoadProcs = new(StringComparer.OrdinalIgnoreCase) { "Form_Load", "Form_Initialize", "MDIForm_Load", "MDIForm_Initialize", "UserControl_Initialize" };
    static readonly HashSet<string> NonReferenceProps = new(StringComparer.OrdinalIgnoreCase) { "Caption", "Text", "ToolTipText", "Tag", "Name", "DataField", "Index" };

    readonly VbProject _project;
    readonly CleanupOptions _options;
    readonly Dictionary<string, List<VbSourceFile>> _formsByName = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<ControlGroup>> _groupsByName = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<ControlGroup>> _handlerOwners = new(StringComparer.OrdinalIgnoreCase);
    readonly List<(ControlNode From, string Prop, ControlGroup To)> _designRefs = new();

    CleanupAnalyzer(VbProject project, CleanupOptions options)
    {
        _project = project;
        _options = options;
    }

    public static AnalysisResult Analyze(VbProject project, CleanupOptions options) =>
        new CleanupAnalyzer(project, options).Run();

    AnalysisResult Run()
    {
        var result = new AnalysisResult { Project = _project, Options = _options };
        result.Warnings.AddRange(_project.Warnings);

        BuildIndexes();
        foreach (var u in _project.Files.SelectMany(f => f.Units)) Precompute(u);
        ComputeStaticReasons(result);

        var allGroups = _project.Files.SelectMany(f => f.Groups.Values).ToList();
        var removed = new HashSet<ControlGroup>();
        for (int iter = 1; iter <= 100; iter++)
        {
            result.Iterations = iter;
            var next = Iterate(allGroups, removed);
            if (next.SetEquals(removed)) break;
            removed = next;
            foreach (var g in allGroups) g.Removed = removed.Contains(g);
        }

        Verify(removed, result);
        BuildResults(result, removed);
        return result;
    }

    void BuildIndexes()
    {
        foreach (var f in _project.Files.Where(f => f.HasDesigner))
        {
            Add(_formsByName, f.ObjectName, f);
            foreach (var g in f.Groups.Values)
            {
                Add(_groupsByName, g.Name, g);
                foreach (var h in g.Handlers) Add(_handlerOwners, h.Name, g);
            }
            foreach (var node in f.Root!.SelfAndDescendants())
                foreach (var (prop, raw) in node.Properties)
                {
                    if (NonReferenceProps.Contains(prop)) continue;
                    var val = PropertyValue.Unquote(raw);
                    if (val != null && f.Groups.TryGetValue(val, out var target) && target != node.Group)
                        _designRefs.Add((node, prop, target));
                }
        }

        static void Add<T>(Dictionary<string, List<T>> d, string key, T value)
        {
            if (!d.TryGetValue(key, out var list)) d[key] = list = new List<T>();
            list.Add(value);
        }
    }

    // ---------------------------------------------------------------- per statement

    void Precompute(CodeUnit u)
    {
        u.ResetAnalysis();
        var f = u.File;
        var toks = u.Line.Tokens;
        int s = u.TokStart, e = u.TokEnd;

        for (int i = s; i < e; i++)
        {
            var t = toks[i];
            if (t.Kind == TokKind.String && _groupsByName.ContainsKey(t.Text.Trim())) u.Strings.Add(t.Text.Trim());
            else if (t.Kind == TokKind.Ident)
            {
                if (t.Is("Controls")) u.UsesControls = true;
                if (u.Proc?.Header != u && _handlerOwners.ContainsKey(t.Text)) u.CalledNames.Add(t.Text);
            }
        }

        AssignMatch? removable = null;
        if (u.IsSingleLineIf)
        {
            int thenIdx = VbLexer.FindIdent(toks, s, e, "Then");
            int elseIdx = VbLexer.FindIdent(toks, thenIdx + 1, e, "Else");
            var thenParts = VbLexer.SplitColons(toks, thenIdx + 1, elseIdx < 0 ? e : elseIdx);
            var elseParts = elseIdx < 0 ? new List<(int, int)>() : VbLexer.SplitColons(toks, elseIdx + 1, e);
            foreach (var (a, b) in thenParts.Concat(elseParts))
            {
                var m = MatchAssign(f, toks, a, b);
                if (m?.Group != null && m.IsVisible)
                    u.VisEffects.Add(new VisEffect(m.Group, m.Kind == TargetKind.VisibleShow, m.Index, m.IndexUnknown, Conditional: true, u));
                if (elseIdx < 0 && thenParts.Count == 1 && m?.Group != null) removable = m;
            }
        }
        else
        {
            var m = MatchAssign(f, toks, s, e);
            if (m?.Group != null)
            {
                if (m.IsVisible)
                    u.VisEffects.Add(new VisEffect(m.Group, m.Kind == TargetKind.VisibleShow, m.Index, m.IndexUnknown, Conditional: u.Depth > 0, u));
                removable = m;
            }
        }

        if (removable != null && f.IsInsideRoot) u.Target = removable;

        for (int i = s; i < e; i++)
        {
            if (u.Target != null && u.Target.Excluded.Contains(i)) continue;
            var refs = ResolveRef(f, toks, s, i, out bool ambiguous);
            if (refs == null) continue;
            var reason = ambiguous
                ? $"Belirsiz nesne üzerinden erişim — {u.Location}: {Snip(u.Text)}"
                : $"Kodda kullanılıyor — {u.Location}: {Snip(u.Text)}";
            foreach (var g in refs) u.BlockRefs.Add(new BlockRef(g, reason));
        }
    }

    static string Snip(string s) => s.Length <= 90 ? s : s[..87] + "...";

    /// <summary>Resolves the identifier at <paramref name="i"/> to the control groups it may refer to.</summary>
    IReadOnlyList<ControlGroup>? ResolveRef(VbSourceFile f, List<Token> toks, int s, int i, out bool ambiguous)
    {
        ambiguous = false;
        var t = toks[i];
        if (t.Kind != TokKind.Ident) return null;

        bool dotted = i - 1 >= s && toks[i - 1].IsMemberAccess;
        if (!dotted)
            return f.HasDesigner && f.Groups.TryGetValue(t.Text, out var own) ? new[] { own } : null;

        if (i - 2 >= s && toks[i - 2].Kind == TokKind.Ident && !(i - 3 >= s && toks[i - 3].IsMemberAccess))
        {
            var q = toks[i - 2].Text;
            if (SelfNames.Contains(q) && f.HasDesigner)
                return f.Groups.TryGetValue(t.Text, out var own) ? new[] { own } : null;
            if (_formsByName.TryGetValue(q, out var forms))
            {
                var gs = forms.Select(x => x.Groups.GetValueOrDefault(t.Text)).OfType<ControlGroup>().ToList();
                if (gs.Count > 0) return gs;
                return null; // a property/method of the form itself
            }
        }

        if (_groupsByName.TryGetValue(t.Text, out var all))
        {
            ambiguous = true;
            return all;
        }
        return null;
    }

    /// <summary>
    /// Matches statements that only touch one control and can be deleted along with it:
    /// <c>[Set] [Me.|frmX.]ctl[(i)][.Prop...] = expr</c> or <c>ctl.SafeMethod args</c>.
    /// </summary>
    AssignMatch? MatchAssign(VbSourceFile f, List<Token> toks, int a, int b)
    {
        var excluded = new HashSet<int>();
        int p = a;
        if (p < b && (toks[p].Is("Set") || toks[p].Is("Let"))) p++;

        IReadOnlyList<ControlGroup>? groups = null;
        int ctrl = -1;
        if (p + 2 < b && toks[p].Kind == TokKind.Ident && toks[p + 1].IsMemberAccess && toks[p + 2].Kind == TokKind.Ident)
        {
            var q = toks[p].Text;
            var name = toks[p + 2].Text;
            if (SelfNames.Contains(q) && f.HasDesigner)
            {
                if (f.Groups.TryGetValue(name, out var g)) { groups = new[] { g }; ctrl = p + 2; }
            }
            else if (_formsByName.TryGetValue(q, out var forms))
            {
                var gs = forms.Select(x => x.Groups.GetValueOrDefault(name)).OfType<ControlGroup>().ToList();
                if (gs.Count > 0) { groups = gs; ctrl = p + 2; }
            }
        }
        if (ctrl < 0 && p < b && toks[p].Kind == TokKind.Ident && f.HasDesigner && f.Groups.TryGetValue(toks[p].Text, out var bare))
        {
            groups = new[] { bare };
            ctrl = p;
        }
        if (ctrl < 0 || groups == null) return null;

        for (int k = a; k <= ctrl; k++) excluded.Add(k);
        p = ctrl + 1;

        int? index = null;
        bool indexUnknown = false;
        if (p < b && toks[p].Kind == TokKind.LParen)
        {
            int r = VbLexer.MatchParen(toks, p, b);
            if (r < 0) return null;
            if (r == p + 2 && toks[p + 1].Kind == TokKind.Number && int.TryParse(toks[p + 1].Text, out var v)) index = v;
            else indexUnknown = true;
            excluded.Add(p);
            excluded.Add(r);
            p = r + 1;
        }

        var props = new List<string>();
        while (p + 1 < b && toks[p].Kind == TokKind.Dot && toks[p + 1].Kind == TokKind.Ident)
        {
            props.Add(toks[p + 1].Text);
            excluded.Add(p);
            excluded.Add(p + 1);
            p += 2;
        }

        if (p < b && toks[p].Kind == TokKind.Equals)
        {
            if (p + 1 >= b) return null;
            excluded.Add(p);
            bool isVisible = props.Count == 1 && props[0].Equals("Visible", StringComparison.OrdinalIgnoreCase);
            var kind = TargetKind.Write;
            if (isVisible)
            {
                bool hide = p + 2 == b && (toks[p + 1].Is("False") || (toks[p + 1].Kind == TokKind.Number && toks[p + 1].Text == "0"));
                kind = hide ? TargetKind.VisibleHide : TargetKind.VisibleShow;
            }
            return new AssignMatch(groups, kind, index, indexUnknown, excluded);
        }

        if (props.Count == 1 && SafeMethods.Contains(props[0]))
            return new AssignMatch(groups, TargetKind.Method, index, indexUnknown, excluded);

        return null;
    }

    // ---------------------------------------------------------------- reasons

    void ComputeStaticReasons(AnalysisResult result)
    {
        foreach (var f in _project.Files)
        {
            if (!f.HasDesigner)
            {
                if (f.Units.Any(u => u.UsesControls))
                    result.Warnings.Add($"{f.DisplayName} 'Controls' koleksiyonunu kullanıyor; formların kontrollerini döngüyle gösteren kod varsa kontrol edin.");
                continue;
            }

            bool controlsUsed = _options.SkipFormsUsingControlsCollection && f.Units.Any(u => u.UsesControls);
            foreach (var g in f.Groups.Values)
            {
                g.StaticReasons.Clear();
                if (!f.IsInsideRoot) g.StaticReasons.Add("Dosya proje klasörünün dışında");
                else if (!f.CanRemoveControls) g.StaticReasons.Add("Bu dosya türünde kontrol silinmiyor");
                if (controlsUsed) g.StaticReasons.Add("Form kodu Controls koleksiyonunu kullanıyor (dinamik erişim olabilir)");

                foreach (var el in g.Elements)
                {
                    if (el.Properties.TryGetValue("Default", out var d) && PropertyValue.IsTrue(d)
                        || el.Properties.TryGetValue("Cancel", out var c) && PropertyValue.IsTrue(c))
                        AddStatic(g, "Default/Cancel özelliği açık (Enter/Esc ile tetiklenebilir)");
                    if (el.IsMenu && el.Properties.ContainsKey("Shortcut"))
                        AddStatic(g, "Menüde kısayol tuşu tanımlı (gizliyken de çalışabilir)");
                }
            }
        }

        foreach (var g in _project.Files.SelectMany(f => f.Groups.Values).Where(g => _options.ForcedKeep.Contains(g.Key)).ToList())
            foreach (var node in g.Elements.SelectMany(e => e.SelfAndDescendants()))
                AddStatic(node.Group!, node.Group == g ? "Kullanıcı tarafından korundu" : $"Kullanıcı tarafından korunan {g.Name} içinde");

        static void AddStatic(ControlGroup g, string reason)
        {
            if (!g.StaticReasons.Contains(reason)) g.StaticReasons.Add(reason);
        }
    }

    HashSet<ControlGroup> Iterate(List<ControlGroup> groups, HashSet<ControlGroup> removed)
    {
        foreach (var g in groups)
        {
            g.Reasons.Clear();
            g.Reasons.AddRange(g.StaticReasons);
            g.Shows.Clear();
            g.LoadHides.Clear();
            g.HasLiveNonVisibleWrite = false;
        }

        var deadProcs = removed.SelectMany(g => g.Handlers).ToHashSet();

        foreach (var u in _project.Files.SelectMany(f => f.Units))
        {
            if (u.Proc != null && deadProcs.Contains(u.Proc)) continue;

            foreach (var v in u.VisEffects)
            {
                if (v.Show) v.Group.Shows.Add(v);
                else if (!v.Conditional && !v.IndexUnknown && !u.AfterExit && u.Depth == 0
                         && u.File == v.Group.File && u.Proc != null && LoadProcs.Contains(u.Proc.Name))
                    v.Group.LoadHides.Add(v);
            }

            if (u.Target?.Group is { } target && removed.Contains(target)) continue; // will be deleted

            var owner = u.Proc?.HandlerOwner;
            foreach (var r in u.BlockRefs)
                if (r.Group != owner) r.Group.AddReason(r.Reason);

            if (u.Target?.Group is { } tg && tg != owner && !u.Target.IsVisible)
                tg.HasLiveNonVisibleWrite = true;

            foreach (var name in u.CalledNames)
                foreach (var g in _handlerOwners[name])
                    if (u.Proc == null || !u.Proc.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                        g.AddReason($"Olay prosedürü {name} koddan çağrılıyor — {u.Location}");

            foreach (var str in u.Strings)
                foreach (var g in _groupsByName[str])
                    g.AddReason($"Adı metin olarak geçiyor (\"{str}\") — {u.Location}");
        }

        foreach (var g in groups)
        {
            if (g.HasLiveNonVisibleWrite && g.Handlers.Any(h => h.HasBody))
                g.AddReason("Kod bu kontrolün değerini değiştiriyor ve olay prosedürleri var (olay tetiklenebilir)");
        }

        foreach (var (from, prop, to) in _designRefs)
            if (from.Group == null || !removed.Contains(from.Group))
                to.AddReason($"{from.DisplayName} kontrolünün '{prop}' özelliği bu kontrole bağlı");

        foreach (var f in _project.Files.Where(f => f.HasDesigner))
            ComputeVisibility(f.Root!);

        while (true)
        {
            foreach (var f in _project.Files.Where(f => f.HasDesigner))
            {
                foreach (var n in f.Root!.Descendants()) { n.WillRemove = false; n.IsRemovalRoot = false; }
                MarkRemoval(f.Root);
            }

            bool changed = false;
            foreach (var g in groups.Where(g => g.Elements.Count > 1))
            {
                int cnt = g.Elements.Count(e => e.WillRemove);
                if (cnt > 0 && cnt < g.Elements.Count)
                {
                    g.AddReason("Kontrol dizisinin bazı elemanları korunuyor (dizi bütün olarak silinebilir)");
                    changed = true;
                }
            }
            if (!changed) break;
        }

        return groups.Where(g => g.Elements.Count > 0 && g.Elements.All(e => e.WillRemove)).ToHashSet();
    }

    void ComputeVisibility(ControlNode root)
    {
        foreach (var n in root.Descendants())
        {
            var g = n.Group!;
            bool shown = g.Shows.Any(v => v.IndexUnknown || v.Index == null || v.Index == n.Index);
            bool hideOnLoad = _options.IncludeHiddenInFormLoad && g.LoadHides.Any(v => v.Index == n.Index);
            n.HiddenForever = !shown && (!n.DesignVisible || hideOnLoad);
            n.HiddenWhy = !n.HiddenForever ? "" : !n.DesignVisible ? "Tasarımda Visible = False" : "Form_Load içinde koşulsuz gizleniyor";

            var parent = n.Parent!;
            n.EffHidden = n.HiddenForever;
            if (!n.EffHidden && !parent.IsRoot && !parent.IsMenu && parent.EffHidden)
            {
                n.EffHidden = true;
                n.HiddenWhy = $"Hiç görünmeyen kapsayıcı içinde: {parent.DisplayName}";
            }
        }
    }

    static void MarkRemoval(ControlNode node)
    {
        foreach (var c in node.Children)
        {
            if (c.EffHidden && SubtreeOk(c))
            {
                c.IsRemovalRoot = true;
                foreach (var d in c.SelfAndDescendants()) d.WillRemove = true;
            }
            else MarkRemoval(c);
        }
    }

    static bool SubtreeOk(ControlNode n) => n.Group!.Reasons.Count == 0 && n.Children.All(SubtreeOk);

    // ---------------------------------------------------------------- results

    void Verify(HashSet<ControlGroup> removed, AnalysisResult result)
    {
        var deadProcs = removed.SelectMany(g => g.Handlers).ToHashSet();
        foreach (var u in _project.Files.SelectMany(f => f.Units))
        {
            if (u.Proc != null && deadProcs.Contains(u.Proc)) continue;
            if (u.Target?.Group is { } tg && removed.Contains(tg)) continue;
            foreach (var r in u.BlockRefs.Where(r => removed.Contains(r.Group) && r.Group != u.Proc?.HandlerOwner))
                result.Warnings.Add($"İç tutarlılık uyarısı: {r.Group} silinecek ama hâlâ kullanılıyor — {u.Location}");
        }
    }

    void BuildResults(AnalysisResult result, HashSet<ControlGroup> removed)
    {
        var deadProcs = removed.SelectMany(g => g.Handlers).ToHashSet();
        foreach (var f in _project.Files)
        {
            var fr = new FileResult { File = f };
            if (f.HasDesigner)
            {
                foreach (var n in f.Root!.Descendants())
                {
                    if (n.IsRemovalRoot) fr.RemovalRoots.Add(n);
                    else if (n.EffHidden && !n.WillRemove) fr.KeptHidden.Add(n);
                }
                if (_options.SkipFormsUsingControlsCollection && f.Units.Any(u => u.UsesControls))
                    fr.Notes.Add("Form kodu Controls koleksiyonunu kullandığı için bu formda kontrol silinmedi.");
            }
            fr.RemovedProcs.AddRange(f.Procs.Where(deadProcs.Contains));
            fr.RemovedUnits.AddRange(f.Units.Where(u =>
                (u.Proc == null || !deadProcs.Contains(u.Proc)) && u.Target?.Group is { } g && removed.Contains(g)));
            result.Files.Add(fr);
        }
    }

    /// <summary>Reasons to show for a hidden control that is kept.</summary>
    public static IReadOnlyList<string> KeepReasons(ControlNode n)
    {
        if (n.Group!.Reasons.Count > 0) return n.Group.Reasons;
        var blocked = n.Descendants().FirstOrDefault(d => d.Group!.Reasons.Count > 0);
        return blocked != null
            ? new[] { $"İçindeki kontrol korunuyor: {blocked.DisplayName} ({blocked.Group!.Reasons[0]})" }
            : new[] { "Bilinmeyen neden" };
    }
}
