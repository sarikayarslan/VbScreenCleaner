using System.Diagnostics;
using System.Text;
using VbScreenCleaner.Core;
using Xunit.Abstractions;

namespace VbScreenCleaner.Tests;

/// <summary>
/// Generates random VB5 projects whose correct outcome is known by construction and compares the
/// analyzer against an independent oracle, then checks invariants on the written output.
/// </summary>
public sealed class RandomizedTests(ITestOutputHelper output)
{
    sealed class GenControl
    {
        public required string Name { get; init; }
        public required bool IsFrame { get; init; }
        public required bool DesignHidden { get; init; }
        public GenControl? Parent { get; init; }
        public List<GenControl> Children { get; } = new();
        public bool Shown, LoadHidden, Read;
    }

    sealed class GenForm
    {
        public required string Name { get; init; }
        public List<GenControl> Controls { get; } = new();
        public List<string> LoadCode { get; } = new();
        public List<string> ClickCode { get; } = new();
    }

    /// <summary><paramref name="ModuleLive"/>: whether anything calls the module's Islem procedure.</summary>
    sealed record GenProject(List<GenForm> Forms, List<string> ModuleCode, bool ModuleLive);

    static GenProject Generate(int seed, int forms, int controlsPerForm)
    {
        var rnd = new Random(seed);
        var project = new GenProject(new List<GenForm>(), new List<string>(), ModuleLive: rnd.Next(2) == 0);
        for (int fi = 0; fi < forms; fi++)
        {
            var form = new GenForm { Name = $"frmG{fi}" };
            var frames = new List<GenControl>();
            for (int k = 0; k < controlsPerForm; k++)
            {
                GenControl? parent = frames.Count > 0 && rnd.NextDouble() < 0.5 ? frames[rnd.Next(frames.Count)] : null;
                int depth = 0;
                for (var p = parent; p != null; p = p.Parent) depth++;
                bool isFrame = depth < 5 && rnd.NextDouble() < 0.25;
                var c = new GenControl
                {
                    Name = (isFrame ? "fra" : "txt") + $"{fi}_{k}",
                    IsFrame = isFrame,
                    DesignHidden = rnd.NextDouble() < 0.45,
                    Parent = parent,
                };
                (parent?.Children ?? form.Controls).Add(c);
                if (isFrame) frames.Add(c);

                for (int op = rnd.Next(4); op > 0; op--) AddOp(rnd, project, form, c);
            }
            project.Forms.Add(form);
        }

        // Join some adjacent plain statements with ':' to exercise statement splitting/rebuilding.
        foreach (var f in project.Forms)
            for (int i = f.LoadCode.Count - 2; i >= 0; i--)
                if (rnd.NextDouble() < 0.2 && !f.LoadCode[i].StartsWith("If") && !f.LoadCode[i + 1].StartsWith("If"))
                {
                    f.LoadCode[i] += ": " + f.LoadCode[i + 1];
                    f.LoadCode.RemoveAt(i + 1);
                }

        // A dead module never runs: its statements must neither show nor keep any control.
        if (project.ModuleLive) project.Forms[0].LoadCode.Insert(0, "Islem");
        return project;
    }

    static void AddOp(Random rnd, GenProject p, GenForm f, GenControl c)
    {
        string own = rnd.Next(3) switch { 0 => "Me." + c.Name, 1 => $"{f.Name}.{c.Name}", _ => c.Name };
        string qual = $"{f.Name}.{c.Name}";
        int where = rnd.Next(3); // 0 Form_Load, 1 cmdGo_Click, 2 module
        bool runs = where != 2 || p.ModuleLive;
        double r = rnd.NextDouble();
        if (r < 0.30) Place($"{(where == 2 ? qual : own)}.Tag = \"w\"");
        else if (r < 0.50)
        {
            // Only Form_Load hides count as "hidden on load"; elsewhere a hide changes nothing.
            if (where == 0) c.LoadHidden = true;
            Place($"{(where == 2 ? qual : own)}.Visible = False");
        }
        else if (r < 0.62)
        {
            c.Shown |= runs;
            Place($"{(where == 2 ? qual : own)}.Visible = True");
        }
        else if (r < 0.74)
        {
            c.Read |= runs;
            Place($"s = {(where == 2 ? qual : own)}.Tag");
        }
        else if (r < 0.88) f.LoadCode.Add($"If gB Then {own}.Visible = False");
        else f.LoadCode.Add($"If gB Then {own}.Tag = \"c\"");

        void Place(string stmt)
        {
            if (where == 0) f.LoadCode.Add(stmt);
            else if (where == 1) f.ClickCode.Add(stmt);
            else p.ModuleCode.Add(stmt);
        }
    }

    static void Write(GenProject p, string dir)
    {
        var vbp = new StringBuilder("Type=Exe\r\n");
        foreach (var f in p.Forms)
        {
            var sb = new StringBuilder();
            sb.Append("VERSION 5.00\r\n").Append($"Begin VB.Form {f.Name} \r\n").Append("   Caption         =   \"G\"\r\n");
            sb.Append("   Begin VB.CommandButton cmdGo \r\n      Caption         =   \"Go\"\r\n   End\r\n");
            foreach (var c in f.Controls) WriteControl(sb, c, 1);
            sb.Append("End\r\n").Append($"Attribute VB_Name = \"{f.Name}\"\r\n").Append("Option Explicit\r\nPrivate s As String\r\n\r\n");
            sb.Append("Private Sub Form_Load()\r\n");
            foreach (var l in f.LoadCode) sb.Append("    ").Append(l).Append("\r\n");
            sb.Append("End Sub\r\n\r\nPrivate Sub cmdGo_Click()\r\n");
            foreach (var l in f.ClickCode) sb.Append("    ").Append(l).Append("\r\n");
            sb.Append("End Sub\r\n");
            Vb.Write(Path.Combine(dir, f.Name + ".frm"), sb.ToString());
            vbp.Append($"Form={f.Name}.frm\r\n");
        }
        var mod = new StringBuilder("Attribute VB_Name = \"modG\"\r\nOption Explicit\r\nPublic gB As Boolean\r\n\r\nPublic Sub Islem()\r\n    Dim s As String\r\n");
        foreach (var l in p.ModuleCode) mod.Append("    ").Append(l).Append("\r\n");
        mod.Append("End Sub\r\n");
        Vb.Write(Path.Combine(dir, "modG.bas"), mod.ToString());
        vbp.Append("Module=modG; modG.bas\r\n");
        Vb.Write(Path.Combine(dir, "G.vbp"), vbp.ToString());
    }

    static void WriteControl(StringBuilder sb, GenControl c, int depth)
    {
        var ind = new string(' ', depth * 3);
        sb.Append($"{ind}Begin {(c.IsFrame ? "VB.Frame" : "VB.TextBox")} {c.Name} \r\n");
        sb.Append($"{ind}   Height          =   300\r\n");
        if (c.DesignHidden) sb.Append($"{ind}   Visible         =   0   'False\r\n");
        foreach (var child in c.Children) WriteControl(sb, child, depth + 1);
        sb.Append($"{ind}End\r\n");
    }

    /// <summary>Independent statement of the rules: what must be removed.</summary>
    static HashSet<string> Oracle(GenProject p)
    {
        var removed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        bool HiddenForever(GenControl c) => !c.Shown && (c.DesignHidden || c.LoadHidden);
        bool Eff(GenControl c) => HiddenForever(c) || (c.Parent is { IsFrame: true } parent && Eff(parent));
        bool SubtreeOk(GenControl c) => !c.Read && c.Children.All(SubtreeOk);
        void Collect(GenControl c) { removed.Add(c.Name); c.Children.ForEach(Collect); }
        void Visit(IEnumerable<GenControl> list)
        {
            foreach (var c in list)
                if (Eff(c) && SubtreeOk(c)) Collect(c);
                else Visit(c.Children);
        }
        foreach (var f in p.Forms) Visit(f.Controls);
        return removed;
    }

    static void CheckProject(GenProject p, string dir, string outDir)
    {
        Write(p, dir);
        var result = CleanupAnalyzer.Analyze(VbProject.Load(dir), new CleanupOptions());
        var actual = result.Files.SelectMany(f => f.RemovalRoots).SelectMany(n => n.SelfAndDescendants())
            .Select(n => n.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var expected = Oracle(p);

        var missing = expected.Except(actual).Order().ToList();
        var extra = actual.Except(expected).Order().ToList();
        Assert.True(missing.Count == 0 && extra.Count == 0,
            $"silinmesi gerekip silinmeyen: [{string.Join(", ", missing)}]  silinmemesi gerekip silinen: [{string.Join(", ", extra)}]");
        Assert.DoesNotContain(result.Warnings, w => w.Contains("İç tutarlılık"));

        CleanupWriter.Write(result, outDir);
        var cleaned = VbProject.Load(outDir);
        // The module goes away exactly when it is dead and touches a removed control.
        bool moduleTouchesRemoved = p.ModuleCode.Any(l => VbLexer.Tokenize(l).Any(t => t.Kind == TokKind.Ident && actual.Contains(t.Text)));
        Assert.Equal(!p.ModuleLive && moduleTouchesRemoved, result.For("modG").DeadCodeProcs.Any(pr => pr.Name == "Islem"));

        foreach (var f in cleaned.Files)
        {
            if (f.HasDesigner) Assert.True(f.CodeStart > 0, $"{f.DisplayName}: designer bozuk");
            foreach (var t in f.LogicalLines.SelectMany(l => l.Tokens))
                Assert.False(t.Kind == TokKind.Ident && actual.Contains(t.Text), $"{f.DisplayName}: silinen {t.Text} hâlâ geçiyor");
            Assert.Equal(f.Procs.Count, f.Procs.Count(pr => pr.EndLine >= 0)); // every Sub still closed
        }
        var second = CleanupAnalyzer.Analyze(cleaned, new CleanupOptions());
        Assert.Equal(0, second.TotalRemovedControls);
        Assert.All(second.Files, f => Assert.False(f.HasChanges));
    }

    public static IEnumerable<object[]> Seeds => Enumerable.Range(1, 60).Select(i => new object[] { i });

    [Theory]
    [MemberData(nameof(Seeds))]
    public void Random_project_matches_oracle(int seed)
    {
        using var tmp = new TempDir();
        CheckProject(Generate(seed, forms: 3, controlsPerForm: 25), tmp.Combine("src"), tmp.Combine("out"));
    }

    [Fact]
    public void Large_project_150_forms_200_controls_is_correct_and_fast()
    {
        using var tmp = new TempDir();
        var p = Generate(seed: 2026, forms: 150, controlsPerForm: 200);
        var sw = Stopwatch.StartNew();
        CheckProject(p, tmp.Combine("src"), tmp.Combine("out"));
        sw.Stop();
        output.WriteLine($"150 form x 200 kontrol: analiz + yazma + doğrulama {sw.Elapsed.TotalSeconds:F1} sn");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(60), $"çok yavaş: {sw.Elapsed}");
    }
}
