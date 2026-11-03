using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

/// <summary>samples/LibVb5: the project lives in frm\, shared code in lib\ (referenced as ..\lib\... from the .vbp).</summary>
public sealed class LibraryTests : IDisposable
{
    static string Frm => Path.Combine(TestPaths.RepoRoot, "samples", "LibVb5", "frm");
    static string Lib => Path.Combine(TestPaths.RepoRoot, "samples", "LibVb5", "lib");

    readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    static AnalysisResult Analyze(bool cleanLib = false, bool withLib = true) =>
        CleanupAnalyzer.Analyze(VbProject.Load(Frm, withLib ? new[] { Lib } : null), new CleanupOptions { CleanLibraries = cleanLib });

    [Fact]
    public void Lib_files_are_loaded_as_library_without_outside_warnings()
    {
        var r = Analyze();
        Assert.Equal(new[] { "frmAna" }, r.Project.Files.Where(f => !f.IsLibrary).Select(f => f.ObjectName));
        Assert.Equal(new[] { @"lib\clsLog.cls", @"lib\frmArama.frm", @"lib\modOrtak.bas" },
            r.Project.Files.Where(f => f.IsLibrary).Select(f => f.DisplayName).Order());
        Assert.DoesNotContain(r.Warnings, w => w.Contains("dışında") || w.Contains("bulunamadı"));
    }

    [Fact]
    public void Read_only_lib_keeps_what_it_touches_and_is_not_changed()
    {
        var r = Analyze();
        Assert.Equal(new[] { "txtTemiz" }, r.RemovedNames("frmAna"));
        Assert.Contains(@"lib\modOrtak.bas", r.Reasons("frmAna", "txtLibYazar")); // write in a file we may not edit
        Assert.Contains(@"lib\modOrtak.bas", r.Reasons("frmAna", "txtLibOkur"));
        Assert.False(r.Node("frmAna", "lblLibGoster").EffHidden);                 // lib procedures always count as live
        Assert.Empty(r.RemovedNames("frmArama"));
        Assert.Contains("Lib klasöründeki dosya", r.Reasons("frmArama", "cmdAramaGizli"));
        Assert.All(r.Files.Where(f => f.File.IsLibrary), f => Assert.False(f.HasChanges, f.File.DisplayName));
    }

    [Fact]
    public void Clean_lib_option_cleans_lib_forms_and_statements()
    {
        var r = Analyze(cleanLib: true);
        Assert.Equal(new[] { "txtLibYazar", "txtTemiz" }, r.RemovedNames("frmAna"));
        Assert.Equal(new[] { "frmAna.txtLibYazar.Text = \"\"" }, r.For("modOrtak").RemovedUnits.Select(u => u.Text));
        Assert.Equal(new[] { "cmdAramaGizli" }, r.RemovedNames("frmArama"));
        Assert.Equal(new[] { "cmdAramaGizli_Click" }, r.For("frmArama").RemovedProcs.Select(p => p.Name));
        Assert.False(r.IsRemoved("frmAna", "txtLibOkur"));                         // still read by the lib
        Assert.False(r.Node("frmAna", "lblLibGoster").EffHidden);                 // GosterAna: lib code stays an entry point
    }

    [Fact]
    public void Output_keeps_frm_and_lib_side_by_side_so_vbp_paths_still_work()
    {
        var r = Analyze();
        var outDir = _tmp.Combine("Temiz");
        CleanupWriter.Write(r, outDir);

        Assert.True(File.Exists(Path.Combine(outDir, CleanupWriter.ReportFileName)));
        foreach (var name in new[] { "modOrtak.bas", "frmArama.frm", "clsLog.cls" })
            Assert.Equal(File.ReadAllBytes(Path.Combine(Lib, name)), File.ReadAllBytes(Path.Combine(outDir, "lib", name)));
        Assert.Equal(File.ReadAllBytes(Path.Combine(Frm, "LibProje.vbp")), File.ReadAllBytes(Path.Combine(outDir, "frm", "LibProje.vbp")));
        Assert.DoesNotContain("txtTemiz", File.ReadAllText(Path.Combine(outDir, "frm", "frmAna.frm")));

        // The cleaned project opens with all its files and has nothing more to clean.
        var cleaned = VbProject.Load(Path.Combine(outDir, "frm"), new[] { Path.Combine(outDir, "lib") });
        Assert.Equal(4, cleaned.Files.Count);
        Assert.DoesNotContain(cleaned.Warnings, w => w.Contains("bulunamadı") || w.Contains("dışında"));
        Assert.Equal(0, CleanupAnalyzer.Analyze(cleaned, new CleanupOptions()).TotalRemovedControls);
    }

    [Fact]
    public void Cleaned_lib_is_written_next_to_frm()
    {
        var r = Analyze(cleanLib: true);
        var outDir = _tmp.Combine("Temiz");
        CleanupWriter.Write(r, outDir);
        var mod = File.ReadAllText(Path.Combine(outDir, "lib", "modOrtak.bas"));
        Assert.DoesNotContain("txtLibYazar", mod);
        Assert.Contains("s = frmAna.txtLibOkur.Text", mod);
        Assert.DoesNotContain("cmdAramaGizli", File.ReadAllText(Path.Combine(outDir, "lib", "frmArama.frm")));
    }

    [Fact]
    public void Without_lib_folder_the_files_are_read_only_outsiders_with_a_hint()
    {
        var r = Analyze(withLib: false);
        Assert.Contains(r.Warnings, w => w.Contains("dışında") && w.Contains("Lib klasörü"));
        Assert.False(r.IsRemoved("frmAna", "txtLibYazar"));

        var outDir = _tmp.Combine("Temiz");
        CleanupWriter.Write(r, outDir);
        Assert.True(File.Exists(Path.Combine(outDir, "LibProje.vbp")));          // old layout: frm contents at the top
        Assert.False(Directory.Exists(Path.Combine(outDir, "lib")));
    }

    [Fact]
    public void Target_inside_lib_is_rejected()
    {
        var r = Analyze();
        var ex = Assert.Throws<InvalidOperationException>(() => CleanupWriter.Write(r, Path.Combine(Lib, "Temiz")));
        Assert.Contains(Lib, ex.Message);
        Assert.False(Directory.Exists(Path.Combine(Lib, "Temiz")));
    }

    [Fact]
    public void Lib_inside_the_project_folder_is_still_a_library()
    {
        var proj = _tmp.Combine("Proje");
        Vb.Write(Path.Combine(proj, "frmT.frm"), Vb.Form("frmT", Vb.Control("VB.TextBox", "txtH", visible: false), ""));
        Vb.Write(Path.Combine(proj, "lib", "Ortak.bas"), Vb.Module("Ortak", "Sub Yaz()\r\n    frmT.txtH.Text = \"a\"\r\nEnd Sub\r\n"));
        Vb.Write(Path.Combine(proj, "P.vbp"), "Type=Exe\r\nForm=frmT.frm\r\nModule=Ortak; lib\\Ortak.bas\r\n");

        var r = CleanupAnalyzer.Analyze(VbProject.Load(proj, new[] { Path.Combine(proj, "lib") }), new CleanupOptions());
        Assert.True(r.For("Ortak").File.IsLibrary);
        Assert.False(r.IsRemoved("frmT", "txtH"));

        var folders = CleanupWriter.OutputFolders(r.Project, _tmp.Combine("Out"), out bool keepsLayout);
        Assert.True(keepsLayout);
        Assert.Equal(new[] { _tmp.Combine("Out") }, folders.Values);                 // lib is copied with the project
    }

    [Fact]
    public void Output_folders_mirror_their_common_parent()
    {
        var r = Analyze();
        var folders = CleanupWriter.OutputFolders(r.Project, @"C:\Hedef", out bool keepsLayout);
        Assert.True(keepsLayout);
        Assert.Equal(@"C:\Hedef\frm", folders[VbProject.NormalizeDir(Frm)]);
        Assert.Equal(@"C:\Hedef\lib", folders[VbProject.NormalizeDir(Lib)]);
    }

    [Fact]
    public void Missing_lib_folder_is_reported()
    {
        var ex = Assert.Throws<DirectoryNotFoundException>(() => VbProject.Load(Frm, new[] { _tmp.Combine("yok") }));
        Assert.Contains("Lib klasörü", ex.Message);
    }
}
