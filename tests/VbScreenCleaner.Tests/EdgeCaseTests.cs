using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

/// <summary>Focused scenarios on tiny generated projects.</summary>
public sealed class EdgeCaseTests : IDisposable
{
    readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    AnalysisResult One(string controls, string code, CleanupOptions? options = null) =>
        Vb.Analyze(_tmp, new Dictionary<string, string> { ["frmT.frm"] = Vb.Form("frmT", controls, code) }, options);

    static readonly string Hidden = Vb.Control("VB.TextBox", "txtH", visible: false);

    [Theory]
    [InlineData("s = Me!txtH.Text")]                       // bang syntax read
    [InlineData("Set c = txtH")]                           // object assignment
    [InlineData("Yaz txtH")]                               // passed as argument
    [InlineData("If txtH.Text = \"\" Then Beep")]          // read in a condition
    [InlineData("s = Me.Controls(\"txtH\").Text")]         // name as string
    [InlineData("txtH.SetFocus")]                          // not a safe method
    [InlineData("For Each c In Me.Controls: Next")]        // Controls collection
    public void Any_other_use_keeps_the_hidden_control(string statement)
    {
        var r = One(Hidden, $"Private Sub Form_Click()\r\n    {statement}\r\nEnd Sub\r\n");
        Assert.False(r.IsRemoved("frmT", "txtH"), statement);
    }

    [Theory]
    [InlineData("txtH.Text = \"a\"")]
    [InlineData("Me!txtH.Text = \"a\"")]
    [InlineData("frmT.txtH.Text = \"a\"")]
    [InlineData("txtH = \"a\"")]
    [InlineData("txtH.Font.Bold = True")]
    [InlineData("txtH.Move 0, 0")]
    [InlineData("txtH.Visible = False")]
    [InlineData("If x Then txtH.Text = s")]
    public void Write_only_statements_are_removed_with_the_control(string statement)
    {
        var r = One(Hidden, $"Private Sub Form_Click()\r\n    {statement}\r\n    Beep\r\nEnd Sub\r\n");
        Assert.True(r.IsRemoved("frmT", "txtH"), statement);
        Assert.Single(r.For("frmT").RemovedUnits);
        Assert.Equal("Private Sub Form_Click()\r\n    Beep\r\nEnd Sub", r.Output("frmT").Split("Option Explicit\r\n")[1].TrimEnd());
    }

    [Theory]
    [InlineData("txtH.Visible = True")]
    [InlineData("txtH.Visible = -1")]
    [InlineData("txtH.Visible = Not bGizle")]
    [InlineData("frmT.txtH.Visible = yetki")]
    [InlineData("If x Then txtH.Visible = False Else txtH.Visible = True")]
    public void Showing_in_any_form_means_not_hidden(string statement)
    {
        var r = One(Hidden, $"Private Sub Form_Click()\r\n    {statement}\r\nEnd Sub\r\n");
        Assert.False(r.Node("frmT", "txtH").EffHidden, statement);
    }

    [Theory]
    [InlineData("lbl.Visible = False", true)]
    [InlineData("lbl.Visible = 0", true)]
    [InlineData("lbl.Visible = vbFalse", true)]
    [InlineData("lbl.Visible = (False)", true)]
    [InlineData("lbl.Visible = ((0))", true)]
    [InlineData("lbl.Visible = (False) Or x", false)]
    [InlineData("If a Then lbl.Visible = False", false)]
    [InlineData("If a Then\r\n        lbl.Visible = False\r\n    End If", false)]
    [InlineData("Do While a\r\n        lbl.Visible = False\r\n    Loop", false)]
    [InlineData("GoTo Son\r\n    lbl.Visible = False\r\nSon:", false)]
    [InlineData("On Error GoTo Hata\r\n    lbl.Visible = False", true)]
    [InlineData("On Error Resume Next\r\n    lbl.Visible = False", true)]
    public void Design_visible_control_is_hidden_only_by_unconditional_FormLoad_code(string body, bool removed)
    {
        var r = One(Vb.Control("VB.Label", "lbl"), $"Private Sub Form_Load()\r\n    {body}\r\nEnd Sub\r\n");
        Assert.Equal(removed, r.IsRemoved("frmT", "lbl"));
    }

    [Fact]
    public void Hiding_in_another_event_does_not_count()
    {
        var r = One(Vb.Control("VB.Label", "lbl"), "Private Sub Form_Click()\r\n    lbl.Visible = False\r\nEnd Sub\r\n");
        Assert.False(r.IsRemoved("frmT", "lbl"));
    }

    [Fact]
    public void Hiding_another_forms_control_in_this_FormLoad_does_not_count()
    {
        var r = Vb.Analyze(_tmp, new Dictionary<string, string>
        {
            ["frmA.frm"] = Vb.Form("frmA", "", "Private Sub Form_Load()\r\n    frmB.lbl.Visible = False\r\nEnd Sub\r\n"),
            ["frmB.frm"] = Vb.Form("frmB", Vb.Control("VB.Label", "lbl"), ""),
        });
        Assert.False(r.IsRemoved("frmB", "lbl"));
    }

    [Fact]
    public void Array_element_index_must_match_for_FormLoad_hide()
    {
        var controls = Vb.Control("VB.Label", "lbl", index: 0) + Vb.Control("VB.Label", "lbl", index: 1);
        var r = One(controls, "Private Sub Form_Load()\r\n    lbl(0).Visible = False\r\n    lbl(1).Visible = False\r\nEnd Sub\r\n");
        Assert.True(r.IsRemoved("frmT", "lbl"));

        using var tmp2 = new TempDir();
        var r2 = Vb.Analyze(tmp2, new Dictionary<string, string>
        {
            ["frmT.frm"] = Vb.Form("frmT", controls, "Private Sub Form_Load()\r\n    lbl(0).Visible = False\r\nEnd Sub\r\n"),
        });
        Assert.False(r2.IsRemoved("frmT", "lbl")); // lbl(1) stays visible, so the array stays
    }

    [Fact]
    public void Visible_child_inside_hidden_menu_is_not_treated_as_hidden()
    {
        var menu = Vb.Control("VB.Menu", "mnuPop", visible: false, children: "      Begin VB.Menu mnuA \r\n      End\r\n");
        var r = One(menu, "Private Sub Form_Click()\r\n    PopupMenu mnuPop\r\nEnd Sub\r\nPrivate Sub mnuA_Click()\r\n    Beep\r\nEnd Sub\r\n");
        Assert.False(r.IsRemoved("frmT", "mnuA"));
        Assert.Empty(r.For("frmT").RemovedProcs);
    }

    [Fact]
    public void Design_property_reference_from_removed_control_does_not_block()
    {
        // Both hidden: the ListView pointing at the TextBox goes too, so the TextBox can go.
        var controls = Vb.Control("VB.TextBox", "txtH", visible: false)
                       + Vb.Control("ComctlLib.ListView", "lv", visible: false, extra: "      Icons           =   \"txtH\"\r\n");
        var r = One(controls, "");
        Assert.True(r.IsRemoved("frmT", "lv"));
        Assert.True(r.IsRemoved("frmT", "txtH"));
    }

    [Fact]
    public void Shared_module_outside_project_folder_blocks_removal()
    {
        var root = _tmp.Combine("Proje");
        Vb.Write(Path.Combine(root, "frmT.frm"), Vb.Form("frmT", Hidden, ""));
        Vb.Write(_tmp.Combine("Ortak", "Ortak.bas"), Vb.Module("Ortak", "Sub X()\r\n    frmT.txtH.Text = \"a\"\r\nEnd Sub\r\n"));
        Vb.Write(Path.Combine(root, "P.vbp"), "Type=Exe\r\nForm=frmT.frm\r\nModule=Ortak; ..\\Ortak\\Ortak.bas\r\n");

        var r = CleanupAnalyzer.Analyze(VbProject.Load(root), new CleanupOptions());
        Assert.False(r.IsRemoved("frmT", "txtH"));
        Assert.Contains(r.Warnings, w => w.Contains("dışında"));
    }

    [Fact]
    public void Without_vbp_all_vb_files_in_folder_are_processed()
    {
        var r = Vb.Analyze(_tmp, new Dictionary<string, string> { ["Alt\\frmT.frm"] = Vb.Form("frmT", Hidden, "") }, writeVbp: false);
        Assert.True(r.IsRemoved("frmT", "txtH"));
        Assert.Contains(r.Warnings, w => w.Contains(".vbp"));
    }

    [Fact]
    public void Module_using_Controls_collection_produces_a_warning()
    {
        var r = Vb.Analyze(_tmp, new Dictionary<string, string>
        {
            ["frmT.frm"] = Vb.Form("frmT", Hidden, ""),
            ["Mod.bas"] = Vb.Module("Mod", "Sub Goster(f As Form)\r\n    Dim c As Control\r\n    For Each c In f.Controls\r\n        c.Visible = True\r\n    Next\r\nEnd Sub\r\n"),
        });
        Assert.Contains(r.Warnings, w => w.Contains("Mod.bas") && w.Contains("Controls"));
    }

    // ------------------------------------------------------------------ dead code

    [Fact]
    public void Uncalled_private_sub_does_not_keep_a_control_and_is_deleted()
    {
        var r = One(Hidden, "Private Sub Eski()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n");
        Assert.True(r.IsRemoved("frmT", "txtH"));
        Assert.Equal(new[] { "Eski" }, r.For("frmT").DeadCodeProcs.Select(p => p.Name));
        Assert.DoesNotContain("Eski", r.Output("frmT"));
    }

    [Fact]
    public void Dead_code_option_off_keeps_the_old_behaviour()
    {
        var o = new CleanupOptions { RemoveDeadCode = false };
        var r = One(Hidden, "Private Sub Eski()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n", o);
        Assert.False(r.IsRemoved("frmT", "txtH"));
        Assert.Empty(r.For("frmT").DeadCodeProcs);
    }

    [Theory]
    [InlineData("Private Sub Form_Load()\r\n    Goster\r\nEnd Sub\r\nPrivate Sub Goster()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]   // called
    [InlineData("Private Sub Form_Load()\r\n    Call Goster\r\nEnd Sub\r\nPrivate Sub Goster()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]
    [InlineData("Private Sub Form_Load()\r\n    x = Hesap()\r\nEnd Sub\r\nPrivate Function Hesap()\r\n    Hesap = txtH.Text\r\nEnd Function\r\n")]
    [InlineData("Private Sub Form_Load()\r\n    CallByName Me, \"Goster\", VbMethod\r\nEnd Sub\r\nPublic Sub Goster()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]
    [InlineData("Private Sub Form_Load()\r\n    p = AddressOf Goster\r\nEnd Sub\r\nPrivate Sub Goster()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]
    [InlineData("Private Sub IArayuz_Yap()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]                                // Implements / WithEvents
    [InlineData("Private Sub Form_Load()\r\n    A\r\nEnd Sub\r\nPrivate Sub A()\r\n    B\r\nEnd Sub\r\nPrivate Sub B()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n")]
    public void Reachable_procedures_still_keep_controls(string code)
    {
        var r = One(Hidden, code);
        Assert.False(r.IsRemoved("frmT", "txtH"));
        Assert.Empty(r.For("frmT").DeadCodeProcs);
    }

    [Fact]
    public void Public_form_procedure_called_from_a_module_is_live_only_if_the_module_procedure_is()
    {
        var files = new Dictionary<string, string>
        {
            ["frmT.frm"] = Vb.Form("frmT", Hidden, "Public Sub Yenile()\r\n    MsgBox txtH.Text\r\nEnd Sub\r\n"),
            ["Mod.bas"] = Vb.Module("Mod", "Sub Main()\r\n    frmT.Yenile\r\nEnd Sub\r\n"),
        };
        Assert.False(Vb.Analyze(_tmp, files).IsRemoved("frmT", "txtH"));       // Sub Main is the start point

        using var tmp2 = new TempDir();
        files["Mod.bas"] = Vb.Module("Mod", "Sub Calistir()\r\n    frmT.Yenile\r\nEnd Sub\r\n");
        var r = Vb.Analyze(tmp2, files);
        Assert.True(r.IsRemoved("frmT", "txtH"));                              // nothing calls Calistir
        Assert.Equal(new[] { "Calistir" }, r.For("Mod").DeadCodeProcs.Select(p => p.Name)); // calls deleted Yenile
        Assert.Equal(new[] { "Yenile" }, r.For("frmT").DeadCodeProcs.Select(p => p.Name));
    }

    [Fact]
    public void Public_class_member_is_an_entry_point()
    {
        var r = Vb.Analyze(_tmp, new Dictionary<string, string>
        {
            ["frmT.frm"] = Vb.Form("frmT", Hidden, ""),
            ["Cls.cls"] = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Cls\"\r\nOption Explicit\r\n"
                          + "Public Function Deger() As String\r\n    Deger = frmT.txtH.Text\r\nEnd Function\r\n",
        });
        Assert.False(r.IsRemoved("frmT", "txtH"));
    }

    [Fact]
    public void Dead_procedure_calling_a_removed_handler_is_deleted_too()
    {
        var r = One(Vb.Control("VB.CommandButton", "cmdH", visible: false),
            "Private Sub cmdH_Click()\r\n    Beep\r\nEnd Sub\r\nPrivate Sub Tetikle()\r\n    cmdH_Click\r\nEnd Sub\r\n");
        Assert.True(r.IsRemoved("frmT", "cmdH"));
        Assert.Equal(new[] { "cmdH_Click" }, r.For("frmT").RemovedProcs.Select(p => p.Name));
        Assert.Equal(new[] { "Tetikle" }, r.For("frmT").DeadCodeProcs.Select(p => p.Name));
        Assert.DoesNotContain(r.Warnings, w => w.Contains("İç tutarlılık"));
    }

    [Fact]
    public void Removal_cascades_through_several_handlers()
    {
        var controls = Vb.Control("VB.CommandButton", "cmd1", visible: false)
                       + Vb.Control("VB.CommandButton", "cmd2", visible: false)
                       + Vb.Control("VB.TextBox", "txt3", visible: false);
        var r = One(controls, """
            Private Sub cmd1_Click()
                MsgBox cmd2.Caption
            End Sub
            Private Sub cmd2_Click()
                MsgBox txt3.Text
            End Sub
            """);
        Assert.True(r.IsRemoved("frmT", "cmd1"));
        Assert.True(r.IsRemoved("frmT", "cmd2"));
        Assert.True(r.IsRemoved("frmT", "txt3"));
        Assert.True(r.Iterations >= 3);
    }
}

public sealed class WriterTests : IDisposable
{
    readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    AnalysisResult SmallProject() => Vb.Analyze(_tmp, new Dictionary<string, string>
    {
        ["Proje\\frmT.frm"] = Vb.Form("frmT", Vb.Control("VB.TextBox", "txtH", visible: false), ""),
    });

    [Fact]
    public void Target_inside_source_is_rejected()
    {
        var r = SmallProject();
        var ex = Assert.Throws<InvalidOperationException>(() => CleanupWriter.Write(r, Path.Combine(r.Project.RootDir, "Proje", "Temiz")));
        Assert.Contains("içinde", ex.Message);
    }

    [Fact]
    public void Non_empty_target_is_rejected()
    {
        var r = SmallProject();
        using var target = new TempDir();
        File.WriteAllText(target.Combine("var.txt"), "x");
        Assert.Throws<InvalidOperationException>(() => CleanupWriter.Write(r, target.Path));
        Assert.Equal(new[] { "var.txt" }, Directory.GetFiles(target.Path).Select(Path.GetFileName));
    }

    [Fact]
    public void Ansi_bytes_frx_and_line_endings_are_preserved()
    {
        var src = _tmp.Combine("Src");
        Directory.CreateDirectory(src);
        var win1254 = System.Text.CodePagesEncodingProvider.Instance.GetEncoding(1254)!;
        var form = Vb.Form("frmT",
            Vb.Control("VB.Label", "lblTr", extra: "      Caption         =   \"Şirket Ğüİıöç\"\r\n")
            + Vb.Control("VB.TextBox", "txtH", visible: false),
            "Private Sub Form_Load()\r\n    lblTr.Caption = \"Çıkış\": txtH.Text = \"ğ\"\r\nEnd Sub\r\n");
        File.WriteAllBytes(Path.Combine(src, "frmT.frm"), win1254.GetBytes(form));
        var frx = new byte[] { 0x6C, 0x74, 0x00, 0x00, 0xFF, 0xFE, 0x0D, 0x0A };
        File.WriteAllBytes(Path.Combine(src, "frmT.frx"), frx);
        File.WriteAllText(Path.Combine(src, "P.vbp"), "Form=frmT.frm\r\n");

        var r = CleanupAnalyzer.Analyze(VbProject.Load(src), new CleanupOptions());
        var outDir = _tmp.Combine("Out");
        CleanupWriter.Write(r, outDir);

        var written = File.ReadAllBytes(Path.Combine(outDir, "frmT.frm"));
        var expected = win1254.GetBytes(form
            .Replace(Vb.Control("VB.TextBox", "txtH", visible: false), "")
            .Replace(": txtH.Text = \"ğ\"", ""));
        Assert.Equal(expected, written);
        Assert.Equal(frx, File.ReadAllBytes(Path.Combine(outDir, "frmT.frx")));
    }
}
