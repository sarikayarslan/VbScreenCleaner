using VbScreenCleaner.Core;

// Usage:
//   VbScreenCleaner.Cli analyze <projectDir> [--no-load-hidden] [--no-skip-controls] [--no-dead-code]
//   VbScreenCleaner.Cli clean   <projectDir> <targetDir> [--no-load-hidden] [--no-skip-controls] [--no-dead-code]
if (args.Length < 2 || (args[0] == "clean" && args.Length < 3))
{
    Console.WriteLine("Kullanım: analyze <proje klasörü> | clean <proje klasörü> <hedef klasör>  [--no-load-hidden] [--no-skip-controls] [--no-dead-code]");
    return 1;
}

var options = new CleanupOptions
{
    IncludeHiddenInFormLoad = !args.Contains("--no-load-hidden"),
    SkipFormsUsingControlsCollection = !args.Contains("--no-skip-controls"),
    RemoveDeadCode = !args.Contains("--no-dead-code"),
};

var project = VbProject.Load(args[1]);
var result = CleanupAnalyzer.Analyze(project, options);

if (args[0] == "clean")
{
    CleanupWriter.Write(result, args[2], new Progress<string>(Console.WriteLine));
    Console.WriteLine($"Rapor: {Path.Combine(args[2], CleanupWriter.ReportFileName)}");
}
else
{
    Console.OutputEncoding = System.Text.Encoding.UTF8;
    Console.WriteLine(CleanupReport.Build(result));
}
return 0;
