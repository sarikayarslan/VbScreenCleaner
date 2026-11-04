using System.Text;

namespace VbOrphanHandlers;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // Headless mode: VbOrphanHandlers.exe <proje> [lib ...] -o rapor.txt
        if (args.Length > 0)
        {
            var inputs = new List<string>();
            string? outPath = null;
            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "-o" || args[i] == "--out") && i + 1 < args.Length) outPath = args[++i];
                else inputs.Add(args[i]);
            }
            var report = OrphanScanner.Scan(inputs);
            File.WriteAllText(outPath ?? "orphan-rapor.txt", report, new UTF8Encoding(true));
            return 0;
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
        return 0;
    }
}
