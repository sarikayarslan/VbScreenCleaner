using VbScreenCleaner.Core;

namespace VbScreenCleaner.Tests;

static class TestPaths
{
    public static string RepoRoot
    {
        get
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "VbScreenCleaner.sln"))) dir = dir.Parent;
            return dir?.FullName ?? throw new InvalidOperationException("VbScreenCleaner.sln bulunamadı");
        }
    }

    public static string SampleDir => Path.Combine(RepoRoot, "samples", "SampleVb5");
}

sealed class TempDir : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VbScreenCleanerTests", Guid.NewGuid().ToString("N"));

    public TempDir() => Directory.CreateDirectory(Path);

    public string Combine(params string[] parts) => System.IO.Path.Combine(new[] { Path }.Concat(parts).ToArray());

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
    }
}

/// <summary>Builds small VB5 projects on disk for focused scenarios.</summary>
static class Vb
{
    public static string Form(string name, string controls, string code) =>
        "VERSION 5.00\r\n"
        + $"Begin VB.Form {name} \r\n"
        + "   Caption         =   \"Test\"\r\n"
        + "   ClientHeight    =   3000\r\n"
        + controls
        + "End\r\n"
        + $"Attribute VB_Name = \"{name}\"\r\n"
        + "Attribute VB_PredeclaredId = True\r\n"
        + "Option Explicit\r\n"
        + code.Replace("\r\n", "\n").Replace("\n", "\r\n");

    public static string Control(string type, string name, bool visible = true, string extra = "", string children = "", int? index = null) =>
        $"   Begin {type} {name} \r\n"
        + (index is null ? "" : $"      Index           =   {index}\r\n")
        + "      Height          =   375\r\n"
        + (visible ? "" : "      Visible         =   0   'False\r\n")
        + extra
        + children
        + "   End\r\n";

    public static string Module(string name, string code) =>
        $"Attribute VB_Name = \"{name}\"\r\nOption Explicit\r\n" + code.Replace("\r\n", "\n").Replace("\n", "\r\n");

    public static void Write(string path, string content)
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content, VbProject.FileEncoding);
    }

    /// <summary>Writes the files plus a .vbp listing them and returns the analysis.</summary>
    public static AnalysisResult Analyze(TempDir dir, IDictionary<string, string> files, CleanupOptions? options = null, bool writeVbp = true)
    {
        foreach (var (name, content) in files) Write(dir.Combine(name), content);
        if (writeVbp)
        {
            var vbp = new List<string> { "Type=Exe" };
            foreach (var name in files.Keys)
                vbp.Add(System.IO.Path.GetExtension(name).ToLowerInvariant() switch
                {
                    ".frm" => $"Form={name}",
                    ".bas" => $"Module={System.IO.Path.GetFileNameWithoutExtension(name)}; {name}",
                    ".cls" => $"Class={System.IO.Path.GetFileNameWithoutExtension(name)}; {name}",
                    _ => $"UserControl={name}",
                });
            Write(dir.Combine("Test.vbp"), string.Join("\r\n", vbp) + "\r\n");
        }
        return CleanupAnalyzer.Analyze(VbProject.Load(dir.Path), options ?? new CleanupOptions());
    }
}

static class ResultExtensions
{
    public static FileResult For(this AnalysisResult r, string objectName) =>
        r.Files.Single(f => f.File.ObjectName.Equals(objectName, StringComparison.OrdinalIgnoreCase));

    public static ControlNode Node(this AnalysisResult r, string form, string control) =>
        r.For(form).File.Root!.Descendants().First(n => n.Name.Equals(control, StringComparison.OrdinalIgnoreCase));

    public static bool IsRemoved(this AnalysisResult r, string form, string control) => r.Node(form, control).WillRemove;

    public static IReadOnlyList<string> RemovedNames(this AnalysisResult r, string form) =>
        r.For(form).RemovalRoots.SelectMany(n => n.SelfAndDescendants()).Select(n => n.DisplayName).OrderBy(n => n).ToList();

    public static string Reasons(this AnalysisResult r, string form, string control) =>
        string.Join(" | ", CleanupAnalyzer.KeepReasons(r.Node(form, control)));

    public static string Output(this AnalysisResult r, string objectName) =>
        string.Join("\r\n", CleanupWriter.ApplyEdits(r.For(objectName)));
}
