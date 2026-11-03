using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

public sealed class SampleProjectFixture
{
    public AnalysisResult Result { get; } = CleanupAnalyzer.Analyze(VbProject.Load(TestPaths.SampleDir), new CleanupOptions());
}

/// <summary>Decisions on samples/SampleVb5 with default options.</summary>
public sealed class SampleProjectTests(SampleProjectFixture fx) : IClassFixture<SampleProjectFixture>
{
    readonly AnalysisResult _r = fx.Result;

    [Fact]
    public void Loads_all_project_files()
    {
        Assert.Equal(new[] { "frmEk", "frmMain", "frmOther", "Module1" },
            _r.Project.Files.Select(f => f.ObjectName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase));
        Assert.DoesNotContain(_r.Warnings, w => w.Contains("İç tutarlılık"));
    }

    [Fact]
    public void FrmMain_removes_exactly_the_expected_controls()
    {
        Assert.Equal(new[]
        {
            "cmdArrH(0)", "cmdArrH(1)", "cmdExtWrite", "cmdHiddenUnused", "fraHidden", "lblContinued", "lblHideOnLoad",
            "lblX", "lvHidden", "mnuAdmin", "mnuOld", "mnuOldA", "txtA", "txtB", "txtHiddenWrite", "txtOnlyInHandler",
        }.OrderBy(n => n), _r.RemovedNames("frmMain"));
    }

    [Fact]
    public void FrmEk_removes_exactly_the_expected_controls()
    {
        Assert.Equal(new[] { "cmdInPic", "cmdSingleIf", "lblAfterIfBlock", "lstHidden", "picHidden", "txtDefault" }.OrderBy(n => n),
            _r.RemovedNames("frmEk"));
    }

    [Fact]
    public void FrmOther_is_skipped_because_it_uses_Controls_collection()
    {
        Assert.Empty(_r.RemovedNames("frmOther"));
        Assert.Contains("Controls", _r.Reasons("frmOther", "cmdGizli"));
        Assert.NotEmpty(_r.For("frmOther").Notes);
    }

    [Theory]
    [InlineData("frmMain", "cmdHiddenUnused", "Tasarımda Visible = False")]
    [InlineData("frmMain", "lblHideOnLoad", "Form_Load içinde koşulsuz gizleniyor")]
    [InlineData("frmMain", "lblContinued", "Form_Load içinde koşulsuz gizleniyor")]   // Me.ctl.Visible = False
    [InlineData("frmMain", "mnuAdmin", "Form_Load içinde koşulsuz gizleniyor")]      // before the Exit Sub
    [InlineData("frmMain", "lblX", "Hiç görünmeyen kapsayıcı içinde: fraHidden2")]
    [InlineData("frmEk", "lblAfterIfBlock", "Form_Load içinde koşulsuz gizleniyor")] // after End If, depth back to 0
    [InlineData("frmEk", "cmdInPic", "Hiç görünmeyen kapsayıcı içinde: picHidden")]  // shown by code, but parent never is
    public void Removed_controls_report_why_they_are_hidden(string form, string control, string why)
    {
        var node = _r.Node(form, control);
        Assert.True(node.WillRemove);
        Assert.Equal(why, node.HiddenWhy);
    }

    [Theory]
    [InlineData("frmMain", "txtData", "Kodda kullanılıyor")]                // value read: s = txtData.Text
    [InlineData("frmMain", "fraHidden2", "İçindeki kontrol korunuyor: txtData")]
    [InlineData("frmMain", "cmdArr", "Kontrol dizisinin bazı elemanları")] // cmdArr(1) is visible
    [InlineData("frmMain", "cmdCalled", "cmdCalled_Click koddan çağrılıyor")]
    [InlineData("frmMain", "txtStr", "metin olarak geçiyor")]
    [InlineData("frmMain", "chkEvt", "olay tetiklenebilir")]
    [InlineData("frmMain", "cmdWith", "With cmdWith")]
    [InlineData("frmMain", "cmdCancel", "Default/Cancel")]
    [InlineData("frmMain", "mnuPopup", "PopupMenu mnuPopup")]
    [InlineData("frmMain", "mnuHiddenShortcut", "kısayol")]
    [InlineData("frmEk", "cmdDyn", "Load cmdDyn(1)")]
    [InlineData("frmEk", "txtBuddy", "updBuddy kontrolünün 'BuddyControl' özelliği")]
    [InlineData("frmEk", "cmdViaVar", "Belirsiz nesne üzerinden erişim")]   // f.cmdViaVar in Module1
    [InlineData("frmEk", "cmdSetFocus", "cmdSetFocus.SetFocus")]
    [InlineData("frmEk", "cmdMultiIf", "Kodda kullanılıyor")]               // single-line If with two statements
    public void Hidden_but_used_controls_are_kept_with_a_reason(string form, string control, string reasonPart)
    {
        var node = _r.Node(form, control);
        Assert.False(node.WillRemove);
        Assert.True(node.EffHidden);
        Assert.Contains(node, _r.For(form).KeptHidden);
        Assert.Contains(reasonPart, _r.Reasons(form, control), StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("frmMain", "cmdShownInLoad")]   // Visible = True in Form_Load
    [InlineData("frmMain", "cmdExternal")]      // frmMain.cmdExternal.Visible = True in Module1
    [InlineData("frmMain", "lblCondHide")]      // If gAdmin Then ... .Visible = False
    [InlineData("frmMain", "lblAfterExit")]     // hidden only after "If ... Then Exit Sub"
    [InlineData("frmMain", "mnuPopA")]          // child of a popup menu
    [InlineData("frmMain", "imgList")]          // no Visible property, used by lvHidden
    [InlineData("frmEk", "txtSelHide")]         // hidden inside Select Case
    [InlineData("frmEk", "lblIfBlock")]         // hidden inside If block
    [InlineData("frmEk", "cmdVarShow")]         // Visible = gAdmin
    [InlineData("frmEk", "cmdOtherShow")]       // shown from another event handler
    public void Controls_that_can_become_visible_are_not_touched(string form, string control)
    {
        var node = _r.Node(form, control);
        Assert.False(node.WillRemove);
        Assert.False(node.EffHidden);
    }

    [Fact]
    public void Hidden_control_used_only_by_a_removed_handler_is_removed_in_a_later_pass()
    {
        Assert.True(_r.IsRemoved("frmMain", "txtOnlyInHandler"));
        Assert.True(_r.Iterations >= 2);
    }

    [Fact]
    public void Removes_event_handlers_of_removed_controls_only()
    {
        Assert.Equal(new[] { "cmdArrH_Click", "cmdHiddenUnused_Click", "mnuOldA_Click" },
            _r.For("frmMain").RemovedProcs.Select(p => p.Name).OrderBy(n => n));
        Assert.Empty(_r.For("frmEk").RemovedProcs);
    }

    [Fact]
    public void Removes_statements_that_only_touch_removed_controls()
    {
        Assert.Equal(new[]
        {
            "lblHideOnLoad.Visible = False",
            "txtHiddenWrite.Text = \"abc\"",
            "txtHiddenWrite.Text = \"a\" &          \"b\"",
            "Me.lblContinued.Visible = False",
            "mnuAdmin.Visible = False",
            "lblHideOnLoad.Caption = \"son\"",
        }, _r.For("frmMain").RemovedUnits.Select(u => u.Text));

        Assert.Equal(new[]
        {
            "lblAfterIfBlock.Visible = False",
            "txtDefault = \"varsayilan\"",
            "lstHidden.Clear",
            "lstHidden.AddItem \"a\"",
            "cmdInPic.Visible = True",
            "If gAdmin Then cmdSingleIf.Caption = \"a\"",
        }, _r.For("frmEk").RemovedUnits.Select(u => u.Text));

        Assert.Equal(new[] { "frmMain.cmdExtWrite.Caption = \"yeni\"" }, _r.For("Module1").RemovedUnits.Select(u => u.Text));
    }

    [Fact]
    public void Rewritten_code_keeps_surviving_statements_labels_and_comments()
    {
        var main = _r.Output("frmMain");
        Assert.Contains("\r\n    Debug.Print \"yukleniyor\"\r\n", main);  // "txtHiddenWrite.Text = ...: Debug.Print ..."
        Assert.Contains("\r\nErrH:\r\n", main);                          // label kept, statement dropped
        Assert.DoesNotContain("artik kullanilmiyor", main);               // comment of a removed line goes with it
        Assert.DoesNotContain("\"b\"\r\n", main);                         // both physical lines of the continued statement
        Assert.Contains("    If Not gAdmin Then Exit Sub\r\n    lblAfterExit.Visible = False", main);
        Assert.Contains("PopupMenu mnuPopup", main);
        Assert.Contains("Private Sub cmdCalled_Click()", main);

        var ek = _r.Output("frmEk");
        Assert.Contains("If gAdmin Then cmdMultiIf.Caption = \"a\": cmdMultiIf.Tag = \"b\"", ek);
        Assert.DoesNotContain("cmdSingleIf", ek);
    }

    [Fact]
    public void Removed_designer_blocks_are_cut_out_cleanly()
    {
        var main = _r.Output("frmMain");
        Assert.DoesNotContain("Begin VB.Frame fraHidden \r\n", main);
        Assert.Contains("Begin VB.Frame fraHidden2 \r\n", main);
        Assert.DoesNotContain("Begin VB.Menu mnuOld ", main);
        Assert.Contains("Begin VB.Menu mnuHiddenShortcut ", main);
        Assert.DoesNotContain("\r\n\r\n\r\n", main); // no double blank lines left behind by removed procedures
    }
}
