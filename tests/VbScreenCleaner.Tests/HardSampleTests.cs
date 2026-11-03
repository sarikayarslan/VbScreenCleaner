using System.Text;
using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

public sealed class HardSampleFixture : IDisposable
{
    public static string Dir => Path.Combine(TestPaths.RepoRoot, "samples", "HardVb5");

    readonly TempDir _tmp = new();
    public AnalysisResult Result { get; }
    public string OutDir { get; }

    public HardSampleFixture()
    {
        Result = CleanupAnalyzer.Analyze(VbProject.Load(Dir), new CleanupOptions());
        OutDir = _tmp.Combine("Temiz");
        CleanupWriter.Write(Result, OutDir);
    }

    public void Dispose() => _tmp.Dispose();
}

/// <summary>samples/HardVb5: tricky but valid VB5 constructs, every control with a known expected outcome.</summary>
public sealed class HardSampleTests(HardSampleFixture fx) : IClassFixture<HardSampleFixture>
{
    static readonly Encoding Win1254 = CodePagesEncodingProvider.Instance.GetEncoding(1254)!;
    readonly AnalysisResult _r = fx.Result;

    /// <summary>Turkish names as the tool sees them (1254 bytes read through Latin1).</summary>
    static string Tr(string s) => VbProject.FileEncoding.GetString(Win1254.GetBytes(s));

    [Fact]
    public void Loads_project_files_once_and_ignores_files_not_in_any_vbp()
    {
        Assert.Equal(new[] { "clsZor", "ctlKutu", "frmHedef", "frmSekme", "frmTr", "frmZor", "mdiAna", "modZor" },
            _r.Project.Files.Select(f => f.ObjectName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase));
        Assert.Equal(2, _r.Project.ProjectFiles.Count);
        Assert.DoesNotContain(_r.Warnings, w => w.Contains("İç tutarlılık"));
    }

    // ------------------------------------------------------------------ frmZor

    [Fact]
    public void FrmZor_removes_exactly_the_expected_controls()
    {
        Assert.Equal(new[]
        {
            "cmdClickOnly", "cmdGap(0)", "cmdGap(3)", "cmdGap(7)", "cmdMultiLineHeader", "cmdOneLiner",
            "fraInner", "fraL1", "fraL3", "lblBuiltin", "lblCaseVariants", "lblHideAfterResumeNext",
            "lblHideParen", "lblHideWithVbFalse", "lblL5b",
            "lblModBang", "lblModWrite", "mnuGizliAlt", "mnuGizliAltAlt", "mnuGizliUst", "mnuSep",
            "picL2", "picL4", "txtAyni", "txtColonString", "txtCommented", "txtDeadSub", "txtInnerUnused", "txtInString",
            "txtL5", "txtLabelComment", "txtModOlu", "txtNumLabel", "txtOluZincir", "txtSelf", "txtThenColon",
        }.OrderBy(n => n), _r.RemovedNames("frmZor"));
    }

    [Theory]
    [InlineData("lblFunc", "SayacArttir")]                    // removing the line would skip a side effect
    [InlineData("lblFuncIf", "KayitVar")]
    [InlineData("picResize", "picResize_Resize")]             // fires without user interaction
    [InlineData("datGizli", "datGizli_Reposition")]           // Data control repositions on load
    [InlineData("fraMixed", "İçindeki kontrol korunuyor: txtMixedRead")]
    [InlineData("txtMixedRead", "s = txtMixedRead.Text")]
    [InlineData("mnuRecent", "Load mnuRecent(1)")]
    [InlineData("cmdCallKw", "cmdCallKw_Click koddan çağrılıyor")]
    [InlineData("cmdValueTrue", "olay tetiklenebilir")]
    [InlineData("fraHost", "Set txtV.Container = fraHost")]
    [InlineData("txtTypeName", "TypeName(txtTypeName)")]
    [InlineData("txtForms", "Belirsiz")]                      // Forms(0).txtForms
    [InlineData("txtScreen", "Belirsiz")]                     // Screen.ActiveForm.txtScreen
    [InlineData("txtWithMe", "Belirsiz")]                     // With Me / .txtWithMe
    [InlineData("txtDefaultRead", "s = txtDefaultRead")]
    [InlineData("txtPrint", "Print #1, txtPrint.Text")]
    [InlineData("txtSelectCase", "Select Case txtSelectCase.Text")]
    [InlineData("txtClsRead", "clsZor.cls")]                  // public class member: callable from outside
    [InlineData("txtCallByName", "s = txtCallByName.Text")]   // procedure is called by name via CallByName
    public void FrmZor_keeps_hidden_controls_that_are_used(string control, string reasonPart)
    {
        var node = _r.Node("frmZor", control);
        Assert.False(node.WillRemove);
        Assert.True(node.EffHidden);
        Assert.Contains(reasonPart, _r.Reasons("frmZor", control));
    }

    [Theory]
    [InlineData("cmdGo")]
    [InlineData("txtV")]
    [InlineData("lblHideLoopLoad")]     // hidden inside a one-line For loop
    [InlineData("lblHideIfdef")]        // hidden inside #If
    [InlineData("cmdModGoster")]        // shown from a module procedure that MDIForm_Load calls
    [InlineData("mnuDosya")]
    [InlineData("mnuAc")]
    public void FrmZor_leaves_controls_that_can_be_visible(string control)
    {
        Assert.False(_r.Node("frmZor", control).EffHidden);
    }

    [Fact]
    public void FrmZor_removes_handlers_including_one_liners_and_continued_headers()
    {
        Assert.Equal(new[] { "cmdClickOnly_Click", "cmdGap_Click", "cmdMultiLineHeader_MouseDown", "cmdOneLiner_Click", "mnuGizliAltAlt_Click" },
            _r.For("frmZor").RemovedProcs.Select(p => p.Name).OrderBy(n => n));
    }

    [Fact]
    public void Dead_procedures_using_removed_controls_are_deleted_with_their_callers()
    {
        // OluA only calls OluB; OluB reads a control that is now removed, so both must go.
        Assert.Equal(new[] { "HicCagrilmayan", "OluA", "OluB" }, _r.For("frmZor").DeadCodeProcs.Select(p => p.Name).OrderBy(n => n));
        Assert.Equal(new[] { "ModulOlu" }, _r.For("modZor").DeadCodeProcs.Select(p => p.Name));

        var o = _r.Output("frmZor");
        Assert.Contains("Private Sub OluAmaZararsiz()", o);   // dead but harmless: left alone
        Assert.Contains("Public Sub DinamikCagrilan()", o);   // called through CallByName
        Assert.DoesNotContain("Sub OluA()", o);
        Assert.DoesNotContain("Sub HicCagrilmayan()", o);
        Assert.DoesNotContain("ModulOlu", _r.Output("modZor"));
        Assert.Contains("Public Sub ModulYaz()", _r.Output("modZor"));
    }

    [Fact]
    public void FrmZor_removes_the_expected_statements()
    {
        Assert.Equal(new[]
        {
            "lblHideAfterResumeNext.Visible = False",
            "If gA Then: txtThenColon.Tag = \"a\"",
            "lblBuiltin.Caption = Format$(Now, \"hh:nn\") & Left$(\"abc\", 1)",
            "txtSelf.Text = txtSelf.Text & \"a\"",
            "txtColonString.Tag = \"x\"",
            "txtLabelComment.Tag = \"\"",
            "txtNumLabel.Tag = \"\"",
            "lblL5b.Visible = True",
            "LBLCASEVARIANTS.Caption = \"x\"",
            "lblHideWithVbFalse.Visible = vbFalse",
            "lblHideParen.Visible = (False)",
        }, _r.For("frmZor").RemovedUnits.Select(u => u.Text));
    }

    [Fact]
    public void FrmZor_output_code_is_rewritten_correctly()
    {
        var o = _r.Output("frmZor");
        Assert.DoesNotContain("If gA Then", o);                   // no orphan block-If header
        Assert.Contains("\r\n    MsgBox \"a: b\"\r\n", o);          // colon inside string is not a separator
        Assert.Contains("\r\nHata: ' temizle\r\n", o);            // label + comment kept
        Assert.Contains("\r\n10\r\n", o);                         // line number kept
        Assert.Contains("\r\n    ' s = txtCommented.Text\r\n    Rem txtCommented.Visible = True\r\n", o);
        Assert.Contains("lblFunc.Caption = SayacArttir()", o);
        Assert.Contains("If KayitVar() Then lblFuncIf.Caption = \"x\"", o);
        Assert.Contains("#If DEBUGMOD Then\r\n    lblHideIfdef.Visible = False\r\n#End If", o);
        Assert.DoesNotContain("cmdOneLiner", o);
        Assert.DoesNotContain("Shift As Integer", o);              // second line of the continued header
        Assert.Contains("Private Sub cmdCallKw_Click()", o);
        Assert.Contains("Private Sub picResize_Resize()", o);
        Assert.Contains("   Begin VB.Frame fraMixed \r\n", o);     // VB's trailing space survives
        Assert.DoesNotContain("Begin VB.Frame fraInner", o);
    }

    // ------------------------------------------------------------------ other files

    [Fact]
    public void SSTab_children_are_removed_and_tab_lists_renumbered()
    {
        Assert.Equal(new[] { "lblDisari", "txtDetayGizli", "txtSekmeGizli" }, _r.RemovedNames("frmSekme"));

        var o = _r.Output("frmSekme");
        Assert.Contains(
            "      Tab(0).ControlEnabled=   -1  'True\r\n"
            + "      Tab(0).Control(0)=   \"txtSekmeGorunur\"\r\n"
            + "      Tab(0).Control(0).Enabled=   0   'False\r\n"
            + "      Tab(0).Control(1)=   \"txtSekme3\"\r\n"
            + "      Tab(0).Control(1).Enabled=   0   'False\r\n"
            + "      Tab(0).ControlCount=   2\r\n"
            + "      TabCaption(1)   =   \"Detay\"\r\n"
            + "      TabPicture(1)   =   \"frmSekme.frx\":001C\r\n"
            + "      Tab(1).ControlEnabled=   0   'False\r\n"
            + "      Tab(1).Control(0)=   \"txtDetay\"\r\n"
            + "      Tab(1).ControlCount=   1\r\n"
            + "      Begin VB.TextBox txtSekmeGorunur \r\n", o);
        Assert.DoesNotContain("Gizli", o);
    }

    [Fact]
    public void Same_control_name_in_two_forms_is_resolved_per_form()
    {
        Assert.True(_r.IsRemoved("frmZor", "txtAyni"));          // only frmSekme shows its own txtAyni
        Assert.False(_r.Node("frmSekme", "txtAyni").EffHidden);
    }

    [Fact]
    public void Form_whose_Controls_are_looped_from_a_module_is_skipped()
    {
        Assert.Empty(_r.RemovedNames("frmHedef"));
        Assert.Contains("Controls", _r.Reasons("frmHedef", "txtHedef"));
    }

    [Fact]
    public void Module_statements_on_removed_controls_are_removed()
    {
        Assert.Equal(new[] { "frmZor.lblModWrite.Caption = \"x\"", "frmZor!lblModBang.Caption = \"y\"" },
            _r.For("modZor").RemovedUnits.Select(u => u.Text));
    }

    [Fact]
    public void MDI_form_and_UserControl_are_cleaned()
    {
        Assert.Equal(new[] { "cmdMdiGizli", "picDurum" }, _r.RemovedNames("mdiAna"));
        Assert.Equal(new[] { "txtIc" }, _r.RemovedNames("ctlKutu"));
        Assert.Contains("lblIcOkunan.Caption", _r.Reasons("ctlKutu", "lblIcOkunan"));
    }

    [Fact]
    public void Turkish_identifiers_are_removed_and_all_other_bytes_kept()
    {
        Assert.Equal(new[] { Tr("txtŞirketAdı") }, _r.RemovedNames("frmTr"));

        var original = File.ReadAllText(Path.Combine(HardSampleFixture.Dir, "Formlar", "frmTr.frm"), Win1254);
        int b = original.IndexOf("   Begin VB.TextBox txtŞirketAdı ", StringComparison.Ordinal);
        int e = original.IndexOf("   End\r\n", b, StringComparison.Ordinal) + "   End\r\n".Length;
        var expected = (original[..b] + original[e..]).Replace("txtŞirketAdı.Text = \"Çağrı Şahin\": ", "");
        Assert.Equal(Win1254.GetBytes(expected), File.ReadAllBytes(Path.Combine(fx.OutDir, "Formlar", "frmTr.frm")));
    }

    // ------------------------------------------------------------------ written project

    [Theory]
    [InlineData(@"Formlar\frmSekme.frx")]
    [InlineData(@"Formlar\frmEski.frm")]   // in the folder but in no .vbp
    [InlineData(@"Formlar\frmHedef.frm")]
    [InlineData(@"Araclar\Arac.vbp")]
    [InlineData(@"Siniflar\clsZor.cls")]
    [InlineData("ZorProje.vbp")]
    public void Untouched_files_are_copied_byte_for_byte(string path)
    {
        Assert.Equal(File.ReadAllBytes(Path.Combine(HardSampleFixture.Dir, path)), File.ReadAllBytes(Path.Combine(fx.OutDir, path)));
    }

    [Fact]
    public void Output_has_no_token_referencing_a_removed_control_in_its_own_scope()
    {
        var cleaned = VbProject.Load(fx.OutDir);
        var removedByForm = _r.Files.Where(f => f.File.HasDesigner).ToDictionary(
            f => f.File.ObjectName,
            f => f.RemovalRoots.SelectMany(n => n.SelfAndDescendants()).Select(n => n.Name).ToHashSet(StringComparer.OrdinalIgnoreCase));
        var hits = new List<string>();
        foreach (var f in cleaned.Files)
        foreach (var line in f.LogicalLines)
        {
            var t = line.Tokens;
            for (int i = 0; i < t.Count; i++)
            {
                if (t[i].Kind != TokKind.Ident) continue;
                // own scope: bare names in a form; qualified frmX.name / frmX!name anywhere
                bool dotted = i > 0 && t[i - 1].IsMemberAccess;
                string? form = !dotted ? (f.HasDesigner ? f.ObjectName : null)
                    : i > 1 && t[i - 2].Kind == TokKind.Ident ? t[i - 2].Text : null;
                if (form != null && removedByForm.TryGetValue(form, out var names) && names.Contains(t[i].Text))
                    hits.Add($"{f.DisplayName}:{line.FirstLine + 1} {t[i].Text}");
            }
        }
        Assert.Empty(hits);
    }

    [Fact]
    public void Output_reparses_and_a_second_run_removes_nothing()
    {
        var second = CleanupAnalyzer.Analyze(VbProject.Load(fx.OutDir), new CleanupOptions());
        Assert.Equal(0, second.TotalRemovedControls);
        Assert.All(second.Files, f => Assert.False(f.HasChanges, f.File.DisplayName));
        Assert.All(second.Files.Where(f => f.File.HasDesigner), f => Assert.True(f.File.CodeStart > 0));
    }
}
