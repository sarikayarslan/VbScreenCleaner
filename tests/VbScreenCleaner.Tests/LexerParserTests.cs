using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

public sealed class LexerTests
{
    [Theory]
    [InlineData("x = 1 ' yorum", "x = 1 ", "' yorum")]
    [InlineData("s = \"it's\" ' c", "s = \"it's\" ", "' c")]       // apostrophe inside a string
    [InlineData("s = \"a\"\"b\"", "s = \"a\"\"b\"", "")]          // escaped quote
    [InlineData("Rem eski kod", "", "Rem eski kod")]
    [InlineData("x = 1: Rem not", "x = 1: ", "Rem not")]
    [InlineData("Remove = 1", "Remove = 1", "")]                   // not a Rem comment
    public void SplitComment_separates_code_and_comment(string line, string code, string comment)
    {
        Assert.Equal((code, comment), VbLexer.SplitComment(line));
    }

    [Fact]
    public void Tokenize_handles_strings_suffixes_bang_and_named_args()
    {
        var t = VbLexer.Tokenize("s$ = Left$(rs!Ad, 2) & \"a\"\"b\": MsgBox Prompt:=x, Buttons:=&H10");
        Assert.Equal(new[]
        {
            (TokKind.Ident, "s"), (TokKind.Equals, "="), (TokKind.Ident, "Left"), (TokKind.LParen, "("),
            (TokKind.Ident, "rs"), (TokKind.Bang, "!"), (TokKind.Ident, "Ad"), (TokKind.Comma, ","),
            (TokKind.Number, "2"), (TokKind.RParen, ")"), (TokKind.Other, "&"), (TokKind.String, "a\"b"),
            (TokKind.Colon, ":"), (TokKind.Ident, "MsgBox"), (TokKind.Ident, "Prompt"), (TokKind.Other, ":="),
            (TokKind.Ident, "x"), (TokKind.Comma, ","), (TokKind.Ident, "Buttons"), (TokKind.Other, ":="),
            (TokKind.Number, "&H10"),
        }, t.Select(x => (x.Kind, x.Text)));
    }

    [Fact]
    public void Tokenize_reads_turkish_letters_in_identifiers()
    {
        // Windows-1254 "şğı" read through Latin1 become þðý; they must stay one identifier.
        var t = VbLexer.Tokenize("txtþðý.Text = 1");
        Assert.Equal("txtþðý", t[0].Text);
        Assert.Equal(TokKind.Dot, t[1].Kind);
    }

    [Fact]
    public void SplitColons_ignores_named_argument_colons()
    {
        var t = VbLexer.Tokenize("a = 1: Foo x:=2: b = 3");
        Assert.Equal(3, VbLexer.SplitColons(t, 0, t.Count).Count);
    }
}

public sealed class ParserTests : IDisposable
{
    readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    VbSourceFile Parse(string fileName, string content)
    {
        var path = _tmp.Combine(fileName);
        Vb.Write(path, content);
        var f = VbProject.LoadFile(path, _tmp.Path);
        VbParser.Parse(f);
        return f;
    }

    [Fact]
    public void Designer_tree_ignores_property_blocks_and_reads_indexes()
    {
        var frame = Vb.Control("VB.Frame", "fra", children:
            "      Begin VB.TextBox txt \r\n"
            + "         BeginProperty Font \r\n"
            + "            Name            =   \"Tahoma\"\r\n"
            + "         EndProperty\r\n"
            + "      End\r\n");
        var f = Parse("frmT.frm", Vb.Form("frmT", frame + Vb.Control("VB.CommandButton", "cmd", index: 0) + Vb.Control("VB.CommandButton", "cmd", visible: false, index: 1), ""));

        Assert.Equal("frmT", f.ObjectName);
        Assert.Equal(new[] { "fra", "txt", "cmd(0)", "cmd(1)" }, f.Root!.Descendants().Select(n => n.DisplayName));
        Assert.Equal("fra", f.Root.Descendants().Single(n => n.Name == "txt").Parent!.Name);
        Assert.False(f.Root.Descendants().Single(n => n.Name == "txt").Properties.ContainsKey("Name"));
        Assert.Equal(2, f.Groups["cmd"].Elements.Count);
        Assert.True(f.Groups["cmd"].Elements[0].DesignVisible);
        Assert.False(f.Groups["cmd"].Elements[1].DesignVisible);
        Assert.StartsWith("Attribute VB_Name", f.Lines[f.CodeStart]);
    }

    [Fact]
    public void Code_lines_join_continuations_and_detect_labels_and_single_line_ifs()
    {
        var f = Parse("frmT.frm", Vb.Form("frmT", Vb.Control("VB.Label", "lbl"), """
            Private Sub Form_Load()
                x = "a" & _
                    "b"
            Hata: lbl.Caption = "1": y = 2
                If a Then lbl.Caption = "x": z = 1 Else z = 2
                If a Then
                End If
            End Sub
            """));

        var cont = f.LogicalLines.Single(l => l.Code.Contains("\"a\""));
        Assert.Equal(1, cont.LastLine - cont.FirstLine);

        var labelled = f.LogicalLines.Single(l => l.Label != null);
        Assert.Equal("Hata", labelled.Label);
        Assert.Equal(new[] { "lbl.Caption = \"1\"", "y = 2" }, labelled.Units.Select(u => u.Text));

        var ifLine = f.LogicalLines.Single(l => l.Code.Contains("Else z"));
        Assert.True(Assert.Single(ifLine.Units).IsSingleLineIf);

        var blockIf = f.LogicalLines.Single(l => l.Code.Trim() == "If a Then");
        Assert.False(Assert.Single(blockIf.Units).IsSingleLineIf);
    }

    [Fact]
    public void Procedures_get_boundaries_depth_exit_state_and_handler_owner()
    {
        var f = Parse("frmT.frm", Vb.Form("frmT", Vb.Control("VB.CommandButton", "cmd_Save"), """
            Private Declare Function GetTickCount Lib "kernel32" () As Long

            Private Sub cmd_Save_Click()
                If a Then
                    b = 1
                End If
                c = 1
                If d Then Exit Sub
                e = 1
            End Sub

            Public Property Get Deger() As Long
            End Property
            """));

        Assert.Equal(new[] { "cmd_Save_Click", "Deger" }, f.Procs.Select(p => p.Name));
        var click = f.Procs[0];
        Assert.Equal("cmd_Save", click.HandlerOwner!.Name);
        Assert.Equal(f.Lines.FindIndex(l => l.Contains("cmd_Save_Click")), click.HeaderLine);
        Assert.Equal("End Sub", f.Lines[click.EndLine].Trim());

        CodeUnit U(string text) => f.Units.Single(u => u.Text == text);
        Assert.Equal(1, U("b = 1").Depth);
        Assert.Equal(0, U("c = 1").Depth);
        Assert.False(U("c = 1").AfterExit);
        Assert.True(U("e = 1").AfterExit);
        Assert.Equal("Property Get", f.Procs[1].Kind);
    }

    [Fact]
    public void Class_header_is_skipped()
    {
        var f = Parse("Cls.cls", "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Cls\"\r\nOption Explicit\r\n");
        Assert.Equal("Cls", f.ObjectName);
        Assert.Equal(4, f.CodeStart);
        Assert.Null(f.Root);
        Assert.DoesNotContain(f.Units, u => u.Text.Contains("MultiUse"));
    }
}
