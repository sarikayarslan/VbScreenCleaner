using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

/// <summary>Runs the full clean on samples/SampleVb5 and checks the written project.</summary>
public sealed class SampleCleanOutputTests : IDisposable
{
    readonly TempDir _tmp = new();
    readonly string _out;
    readonly AnalysisResult _result;

    public SampleCleanOutputTests()
    {
        _out = _tmp.Combine("Temiz");
        _result = CleanupAnalyzer.Analyze(VbProject.Load(TestPaths.SampleDir), new CleanupOptions());
        CleanupWriter.Write(_result, _out);
    }

    public void Dispose() => _tmp.Dispose();

    [Fact]
    public void Copies_every_source_file_and_adds_the_report()
    {
        var src = Directory.GetFiles(TestPaths.SampleDir).Select(Path.GetFileName).Order();
        var dst = Directory.GetFiles(_out).Select(Path.GetFileName).Where(n => n != CleanupWriter.ReportFileName).Order();
        Assert.Equal(src, dst);
        Assert.True(File.Exists(Path.Combine(_out, CleanupWriter.ReportFileName)));
    }

    [Theory]
    [InlineData("Proje.vbp")]
    [InlineData("frmOther.frm")]
    public void Unchanged_files_are_byte_identical(string name)
    {
        Assert.Equal(File.ReadAllBytes(Path.Combine(TestPaths.SampleDir, name)), File.ReadAllBytes(Path.Combine(_out, name)));
    }

    [Fact]
    public void Output_forms_still_parse_and_contain_exactly_the_kept_controls()
    {
        var cleaned = VbProject.Load(_out);
        foreach (var fr in _result.Files.Where(f => f.File.HasDesigner))
        {
            var after = cleaned.Files.Single(f => f.ObjectName == fr.File.ObjectName);
            Assert.NotNull(after.Root);
            Assert.True(after.CodeStart > 0, $"{after.DisplayName}: designer bloğu kapanmadı");

            var expected = fr.File.Root!.Descendants().Where(n => !n.WillRemove).Select(n => n.DisplayName);
            Assert.Equal(expected, after.Root!.Descendants().Select(n => n.DisplayName));
        }
    }

    [Fact]
    public void Output_code_has_no_reference_to_removed_controls()
    {
        var removed = _result.Files.SelectMany(f => f.RemovalRoots).SelectMany(n => n.SelfAndDescendants())
            .Select(n => n.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var cleaned = VbProject.Load(_out);
        var hits = cleaned.Files.SelectMany(f => f.LogicalLines.SelectMany(l => l.Tokens
                .Where(t => (t.Kind == TokKind.Ident || t.Kind == TokKind.String) && removed.Contains(t.Text))
                .Select(t => $"{f.DisplayName}:{l.FirstLine + 1} {t.Text}")))
            .ToList();
        Assert.Empty(hits);
    }

    [Fact]
    public void Cleaning_the_output_again_removes_nothing_more()
    {
        var second = CleanupAnalyzer.Analyze(VbProject.Load(_out), new CleanupOptions());
        Assert.Equal(0, second.TotalRemovedControls);
        Assert.All(second.Files, f => Assert.False(f.HasChanges, f.File.DisplayName));
    }

    [Fact]
    public void Report_lists_totals_and_decisions()
    {
        var report = File.ReadAllText(Path.Combine(_out, CleanupWriter.ReportFileName));
        Assert.Contains($"Silinen kontrol sayısı: {_result.TotalRemovedControls}", report);
        Assert.Contains("cmdHiddenUnused [VB.CommandButton]", report);
        Assert.Contains("GİZLİ AMA KORUNAN KONTROLLER", report);
        Assert.Contains("txtData", report);
    }
}

/// <summary>Option switches on samples/SampleVb5.</summary>
public sealed class SampleOptionTests
{
    static AnalysisResult Run(Action<CleanupOptions> configure)
    {
        var o = new CleanupOptions();
        configure(o);
        return CleanupAnalyzer.Analyze(VbProject.Load(TestPaths.SampleDir), o);
    }

    [Fact]
    public void Without_FormLoad_rule_design_visible_controls_stay()
    {
        var r = Run(o => o.IncludeHiddenInFormLoad = false);
        Assert.False(r.IsRemoved("frmMain", "lblHideOnLoad"));
        Assert.False(r.IsRemoved("frmMain", "lblContinued"));
        Assert.False(r.IsRemoved("frmMain", "mnuAdmin"));
        Assert.False(r.IsRemoved("frmEk", "lblAfterIfBlock"));
        Assert.True(r.IsRemoved("frmMain", "cmdHiddenUnused")); // design-hidden ones are still removed
    }

    [Fact]
    public void Without_Controls_rule_forms_with_control_loops_are_cleaned()
    {
        var r = Run(o => o.SkipFormsUsingControlsCollection = false);
        Assert.True(r.IsRemoved("frmOther", "cmdGizli"));
    }

    [Fact]
    public void Forced_keep_of_a_container_keeps_its_children()
    {
        var file = VbProject.Load(TestPaths.SampleDir).Files.Single(f => f.ObjectName == "frmMain");
        var r = Run(o => o.ForcedKeep.Add(ControlGroup.MakeKey(file, "fraHidden")));
        Assert.False(r.IsRemoved("frmMain", "fraHidden"));
        Assert.False(r.IsRemoved("frmMain", "txtA"));
        Assert.False(r.IsRemoved("frmMain", "txtB"));
        Assert.Contains("Kullanıcı tarafından korunan fraHidden", r.Reasons("frmMain", "txtA"));
    }

    [Fact]
    public void Forced_keep_of_a_control_keeps_what_its_handler_uses()
    {
        var file = VbProject.Load(TestPaths.SampleDir).Files.Single(f => f.ObjectName == "frmMain");
        var r = Run(o => o.ForcedKeep.Add(ControlGroup.MakeKey(file, "cmdHiddenUnused")));
        Assert.False(r.IsRemoved("frmMain", "cmdHiddenUnused"));
        Assert.False(r.IsRemoved("frmMain", "txtOnlyInHandler"));
        Assert.DoesNotContain(r.For("frmMain").RemovedProcs, p => p.Name == "cmdHiddenUnused_Click");
    }
}
