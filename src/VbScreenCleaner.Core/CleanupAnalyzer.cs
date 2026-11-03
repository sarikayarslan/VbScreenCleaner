namespace VbScreenCleaner.Core;

public sealed class CleanupOptions
{
    /// <summary>Also remove controls that are visible at design time but hidden unconditionally in Form_Load.</summary>
    public bool IncludeHiddenInFormLoad { get; set; } = true;

    /// <summary>Leave forms alone whose code touches the Controls collection (dynamic access by name/loop).</summary>
    public bool SkipFormsUsingControlsCollection { get; set; } = true;

    /// <summary>
    /// Treat procedures nothing calls as dead: their code does not keep controls alive, and they are
    /// deleted when they use a deleted control (or call a deleted procedure).
    /// </summary>
    public bool RemoveDeadCode { get; set; } = true;

    /// <summary>Group keys (<see cref="ControlGroup.Key"/>) the user chose to keep.</summary>
    public HashSet<string> ForcedKeep { get; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class FileResult
{
    public required VbSourceFile File { get; init; }
    public List<ControlNode> RemovalRoots { get; } = new();
    public List<ControlNode> KeptHidden { get; } = new();
    /// <summary>Event procedures of removed controls.</summary>
    public List<ProcInfo> RemovedProcs { get; } = new();
    /// <summary>Procedures nothing calls that had to go because they used removed controls.</summary>
    public List<ProcInfo> DeadCodeProcs { get; } = new();
    public List<CodeUnit> RemovedUnits { get; } = new();
    /// <summary>SSTab-style list entries (owner, property name such as "Tab(0).Control(1)") to drop.</summary>
    public List<(ControlNode Owner, string Property)> TabListRemovals { get; } = new();
    public List<string> Notes { get; } = new();

    public int RemovedControlCount => RemovalRoots.Sum(r => r.SelfAndDescendants().Count());
    public bool HasChanges => RemovalRoots.Count > 0 || RemovedProcs.Count > 0 || DeadCodeProcs.Count > 0 || RemovedUnits.Count > 0;
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

    /// <summary>Events that need the user to interact with the control; a hidden control never gets them.</summary>
    static readonly HashSet<string> InteractiveEvents = new(StringComparer.OrdinalIgnoreCase)
    {
        "Click", "DblClick", "MouseDown", "MouseUp", "MouseMove", "KeyDown", "KeyUp", "KeyPress",
        "GotFocus", "LostFocus", "Validate", "Change", "Scroll", "ItemCheck", "DropDown",
        "DragDrop", "DragOver", "OLEDragDrop", "OLEDragOver", "OLEStartDrag", "OLEGiveFeedback", "OLESetData", "OLECompleteDrag",
        "ColumnClick", "ItemClick", "NodeClick", "ButtonClick", "ButtonMenuClick", "Expand", "Collapse",
        "BeforeLabelEdit", "AfterLabelEdit",
    };

    readonly VbProject _project;
    readonly CleanupOptions _options;
    readonly Dictionary<string, List<VbSourceFile>> _formsByName = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<ControlGroup>> _groupsByName = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<ControlGroup>> _handlerOwners = new(StringComparer.OrdinalIgnoreCase);
    readonly List<(ControlNode From, string Prop, ControlGroup To)> _designRefs = new();
    /// <summary>Functions and Property Gets defined in the project: calling them may have side effects.</summary>
    readonly HashSet<string> _userFunctions = new(StringComparer.OrdinalIgnoreCase);
    HashSet<VbSourceFile> _formsUsingControls = new();

    // Call graph: procedures by name, and the names each procedure (or module-level code) mentions.
    readonly Dictionary<string, List<ProcInfo>> _procsByName = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<ProcInfo, List<string>> _procRefs = new();
    readonly List<(VbSourceFile File, string Name)> _moduleLevelRefs = new();
    /// <summary>SSTab entries "Tab(n).Control(k)" pointing at a control; rewritten instead of blocking.</summary>
    readonly List<(ControlNode Owner, string Prop, ControlGroup To)> _tabRefs = new();
    static readonly System.Text.RegularExpressions.Regex TabControlRx = new(@"^Tab\(\d+\)\.Control\(\d+\)$");

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

        var deleted = ProcsToDelete(removed, DeadProcs(removed));
        Verify(removed, deleted, result);
        BuildResults(result, removed, deleted);
        return result;
    }

    void BuildIndexes()
    {
        foreach (var proc in _project.Files.SelectMany(f => f.Procs))
        {
            if (proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase) || proc.Kind.Equals("Property Get", StringComparison.OrdinalIgnoreCase))
                _userFunctions.Add(proc.Name);
            if (!_procsByName.TryGetValue(proc.Name, out var list)) _procsByName[proc.Name] = list = new List<ProcInfo>();
            list.Add(proc);
        }

        // Every identifier or string that names a procedure is a potential call (CallByName, AddressOf, frmX.Proc ...).
        foreach (var f in _project.Files)
            foreach (var u in f.Units)
                foreach (var t in u.Line.Tokens.Skip(u.TokStart).Take(u.TokEnd - u.TokStart))
                {
                    if (t.Kind is not (TokKind.Ident or TokKind.String) || !_procsByName.ContainsKey(t.Text)) continue;
                    if (u.Proc == null) _moduleLevelRefs.Add((f, t.Text));
                    else
                    {
                        if (!_procRefs.TryGetValue(u.Proc, out var refs)) _procRefs[u.Proc] = refs = new List<string>();
                        refs.Add(t.Text);
                    }
                }

        foreach (var f in _project.Files.Where(f => f.HasDesigner))
        {
            Add(_formsByName, f.ObjectName, f);
            foreach (var g in f.Groups.Values)
            {
                Add(_groupsByName, g.Name, g);
                foreach (var h in g.Handlers) Add(_handlerOwners, h.Name, g);
            }
            foreach (var node in f.Root!.SelfAndDescendants())
                foreach (var (prop, raw) in node.Properties.Select(p => (p.Key, p.Value)).Concat(node.NestedProperties))
                {
                    if (NonReferenceProps.Contains(prop)) continue;
                    var val = PropertyValue.Unquote(raw);
                    if (val == null || !f.Groups.TryGetValue(val, out var target) || target == node.Group) continue;
                    if (TabControlRx.IsMatch(prop) && node.PropertyLines.ContainsKey(prop)) _tabRefs.Add((node, prop, target));
                    else _designRefs.Add((node, prop, target));
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
                if (t.Is("Controls")) ResolveControlsCollection(u, toks, s, i);
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

        if (removable != null && f.IsInsideRoot)
        {
            // Deleting the statement would also skip the call, so the line has to stay.
            var call = Enumerable.Range(s, e - s)
                .Where(i => !removable.Excluded.Contains(i) && toks[i].Kind == TokKind.Ident && _userFunctions.Contains(toks[i].Text))
                .Select(i => toks[i].Text).FirstOrDefault();
            if (call == null) u.Target = removable;
            else
                foreach (var g in removable.Groups)
                    u.BlockRefs.Add(new BlockRef(g, $"Satır {call} fonksiyonunu çağırıyor, silinirse yan etkisi kaybolur — {u.Location}: {Snip(u.Text)}"));
        }

        for (int i = s; i < e; i++)
        {
            if (u.Target != null && u.Target.Excluded.Contains(i)) continue;
            var refs = ResolveRef(f, toks, s, i, out bool ambiguous);
            if (refs == null) continue;
            var reason = ambiguous
                ? $"Belirsiz nesne üzerinden erişim — {u.Location}: {Snip(u.Text)}"
                : $"Kodda kullanılıyor — {u.Location}: {Snip(u.Text)}";
            foreach (var g in refs)
                if (g != u.Target?.Group) // "txt.Text = txt.Text & x" only touches txt itself
                    u.BlockRefs.Add(new BlockRef(g, reason));
        }
    }

    /// <summary>Records whose Controls collection the token at <paramref name="i"/> refers to.</summary>
    void ResolveControlsCollection(CodeUnit u, List<Token> toks, int s, int i)
    {
        var f = u.File;
        bool dotted = i - 1 >= s && toks[i - 1].IsMemberAccess;
        if (!dotted)
        {
            if (f.HasDesigner) u.ControlsForms.Add(f);
            return;
        }
        if (i - 2 >= s && toks[i - 2].Kind == TokKind.Ident && !(i - 3 >= s && toks[i - 3].IsMemberAccess))
        {
            var q = toks[i - 2].Text;
            if (SelfNames.Contains(q) && f.HasDesigner) { u.ControlsForms.Add(f); return; }
            if (_formsByName.TryGetValue(q, out var forms)) { u.ControlsForms.AddRange(forms); return; }
        }
        u.ControlsUnknown = true;
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
                // False, 0, vbFalse — also wrapped in parentheses: (False)
                int ra = p + 1, rb = b;
                while (rb - ra >= 3 && toks[ra].Kind == TokKind.LParen && VbLexer.MatchParen(toks, ra, rb) == rb - 1) { ra++; rb--; }
                bool hide = rb - ra == 1
                    && (toks[ra].Is("False") || toks[ra].Is("vbFalse") || (toks[ra].Kind == TokKind.Number && toks[ra].Text == "0"));
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
        var allUnits = _project.Files.SelectMany(f => f.Units).ToList();
        _formsUsingControls = allUnits.SelectMany(u => u.ControlsForms).ToHashSet();
        foreach (var u in allUnits.Where(u => u.ControlsUnknown).GroupBy(u => u.File).Select(g => g.First()))
            result.Warnings.Add($"{u.File.DisplayName} hangi forma ait olduğu belirsiz bir 'Controls' koleksiyonu kullanıyor ({u.Location}); "
                                + "formların kontrollerini döngüyle gösteren kod varsa kontrol edin.");

        foreach (var f in _project.Files)
        {
            if (!f.HasDesigner) continue;

            bool controlsUsed = _options.SkipFormsUsingControlsCollection && _formsUsingControls.Contains(f);
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

                foreach (var h in g.Handlers.Where(h => h.HasBody))
                    if (!InteractiveEvents.Contains(h.Name[(h.Name.LastIndexOf('_') + 1)..]))
                        AddStatic(g, $"{h.Name} olayı kullanıcı etkileşimi olmadan da tetiklenebilir");
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

        var deadProcs = DeadProcs(removed);

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

    // ---------------------------------------------------------------- dead code

    /// <summary>Procedures whose code can never run: handlers of removed controls and, optionally, unreachable ones.</summary>
    HashSet<ProcInfo> DeadProcs(HashSet<ControlGroup> removed)
    {
        var handlerDead = removed.SelectMany(g => g.Handlers).ToHashSet();
        if (!_options.RemoveDeadCode) return handlerDead;

        var reachable = new HashSet<ProcInfo>();
        var queue = new Queue<ProcInfo>();
        void Mark(ProcInfo p)
        {
            if (!handlerDead.Contains(p) && reachable.Add(p)) queue.Enqueue(p);
        }

        var all = _project.Files.SelectMany(f => f.Procs).ToList();
        foreach (var p in all.Where(IsEntryPoint)) Mark(p);
        foreach (var (file, name) in _moduleLevelRefs)
            foreach (var q in ResolveProc(name, file)) Mark(q);
        while (queue.Count > 0)
        {
            var p = queue.Dequeue();
            if (!_procRefs.TryGetValue(p, out var names)) continue;
            foreach (var name in names)
                foreach (var q in ResolveProc(name, p.File))
                    Mark(q);
        }
        return all.Where(p => !reachable.Contains(p)).ToHashSet();
    }

    /// <summary>Procedures that can run without being called from project code.</summary>
    static bool IsEntryPoint(ProcInfo p)
    {
        if (!p.File.IsInsideRoot) return true;                // shared file: other projects may call it
        if (p.Name.Contains('_')) return true;                // events: Form_Load, cmd_Click, WithEvents, Implements
        if (p.File.Kind == VbFileKind.Module) return p.Name.Equals("Main", StringComparison.OrdinalIgnoreCase);
        if (p.File.Kind is VbFileKind.Class or VbFileKind.UserControl or VbFileKind.PropertyPage
            or VbFileKind.UserDocument or VbFileKind.Designer) return !p.IsPrivate; // object members, callable from outside
        return false;                                         // form procedure: only through a call
    }

    IEnumerable<ProcInfo> ResolveProc(string name, VbSourceFile from) =>
        _procsByName.TryGetValue(name, out var list) ? list.Where(q => q.File == from || !q.IsPrivate) : Enumerable.Empty<ProcInfo>();

    /// <summary>Dead procedures that must be deleted: they use a removed control or call a deleted procedure.</summary>
    HashSet<ProcInfo> ProcsToDelete(HashSet<ControlGroup> removed, HashSet<ProcInfo> dead)
    {
        var delete = removed.SelectMany(g => g.Handlers).ToHashSet();
        foreach (var p in dead.Where(p => p.File.IsInsideRoot && p.EndLine >= 0))
            if (p.Units.Any(u => u.BlockRefs.Any(r => removed.Contains(r.Group))
                                 || (u.Target != null && u.Target.Groups.Any(removed.Contains))
                                 || u.VisEffects.Any(v => removed.Contains(v.Group))))
                delete.Add(p);

        bool changed = true;
        while (changed)
        {
            changed = false;
            foreach (var p in dead.Where(p => !delete.Contains(p) && p.File.IsInsideRoot && p.EndLine >= 0))
                if (_procRefs.TryGetValue(p, out var names)
                    && names.Any(n => ResolveProc(n, p.File).Any(q => q != p && delete.Contains(q))))
                    changed = delete.Add(p) || changed;
        }
        return delete;
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

    /// <summary>Checks that no code surviving the cleanup still names a removed control or a deleted procedure.</summary>
    void Verify(HashSet<ControlGroup> removed, HashSet<ProcInfo> deleted, AnalysisResult result)
    {
        foreach (var u in _project.Files.SelectMany(f => f.Units))
        {
            if (u.Proc != null && deleted.Contains(u.Proc)) continue;
            if (u.Target?.Group is { } tg && removed.Contains(tg)) continue;
            foreach (var r in u.BlockRefs.Where(r => removed.Contains(r.Group)))
                result.Warnings.Add($"İç tutarlılık uyarısı: {r.Group} silinecek ama hâlâ kullanılıyor — {u.Location}");
            foreach (var name in u.Line.Tokens.Skip(u.TokStart).Take(u.TokEnd - u.TokStart)
                         .Where(t => t.Kind == TokKind.Ident).Select(t => t.Text))
                if (ResolveProc(name, u.File).Any(q => deleted.Contains(q) && q != u.Proc) && u.Proc?.Header != u)
                    result.Warnings.Add($"İç tutarlılık uyarısı: silinecek {name} prosedürü hâlâ çağrılıyor — {u.Location}");
        }
    }

    void BuildResults(AnalysisResult result, HashSet<ControlGroup> removed, HashSet<ProcInfo> deleted)
    {
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
                if (_options.SkipFormsUsingControlsCollection && _formsUsingControls.Contains(f))
                    fr.Notes.Add("Form kodu Controls koleksiyonunu kullandığı için bu formda kontrol silinmedi.");
            }
            fr.RemovedProcs.AddRange(f.Procs.Where(p => deleted.Contains(p) && p.HandlerOwner is { } o && removed.Contains(o)));
            fr.DeadCodeProcs.AddRange(f.Procs.Where(p => deleted.Contains(p) && !fr.RemovedProcs.Contains(p)));
            fr.RemovedUnits.AddRange(f.Units.Where(u =>
                (u.Proc == null || !deleted.Contains(u.Proc)) && u.Target?.Group is { } g && removed.Contains(g)));
            result.Files.Add(fr);
        }

        foreach (var (owner, prop, to) in _tabRefs)
            if (removed.Contains(to) && !owner.WillRemove)
                result.Files.Single(fr => fr.File == to.File).TabListRemovals.Add((owner, prop));
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
