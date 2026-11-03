using VbScreenCleaner.Core;

// Usage:
//   VbScreenCleaner.Cli analyze <projectDir>             [--lib <dir>]... [options]
//   VbScreenCleaner.Cli clean   <projectDir> <targetDir> [--lib <dir>]... [options]
// Options: --clean-lib --no-load-hidden --no-skip-controls --no-dead-code
var positional = new List<string>();
var libs = new List<string>();
var flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
for (int i = 0; i < args.Length; i++)
{
    if (args[i].Equals("--lib", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length) libs.Add(args[++i]);
    else if (args[i].StartsWith("--")) flags.Add(args[i]);
    else positional.Add(args[i]);
}

if (positional.Count < 2 || (positional[0] == "clean" && positional.Count < 3))
{
    Console.WriteLine("Kullanım: analyze <proje klasörü> | clean <proje klasörü> <hedef klasör>");
    Console.WriteLine("  [--lib <lib klasörü>]... [--clean-lib] [--no-load-hidden] [--no-skip-controls] [--no-dead-code]");
    return 1;
}

var options = new CleanupOptions
{
    IncludeHiddenInFormLoad = !flags.Contains("--no-load-hidden"),
    SkipFormsUsingControlsCollection = !flags.Contains("--no-skip-controls"),
    RemoveDeadCode = !flags.Contains("--no-dead-code"),
    CleanLibraries = flags.Contains("--clean-lib"),
};

var project = VbProject.Load(positional[1], libs);
var result = CleanupAnalyzer.Analyze(project, options);

if (positional[0] == "clean")
{
    CleanupWriter.Write(result, positional[2], new Progress<string>(Console.WriteLine));
    Console.WriteLine($"Rapor: {Path.Combine(positional[2], CleanupWriter.ReportFileName)}");
}
else
{
    Console.OutputEncoding = System.Text.Encoding.UTF8;
    Console.WriteLine(CleanupReport.Build(result));
}
return 0;
