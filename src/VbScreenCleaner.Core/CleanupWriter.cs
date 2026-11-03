using System.Text;

namespace VbScreenCleaner.Core;

public static class CleanupWriter
{
    public const string ReportFileName = "VbScreenCleaner_Rapor.txt";

    /// <summary>Copies the project to <paramref name="targetDir"/> and writes the cleaned files there.</summary>
    public static void Write(AnalysisResult result, string targetDir, IProgress<string>? progress = null)
    {
        var source = result.Project.RootDir;
        var target = Path.GetFullPath(targetDir).TrimEnd(Path.DirectorySeparatorChar);

        if (target.Equals(source, StringComparison.OrdinalIgnoreCase)
            || target.StartsWith(source + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Hedef klasör kaynak proje klasörünün içinde olamaz.");
        if (Directory.Exists(target) && Directory.EnumerateFileSystemEntries(target).Any())
            throw new InvalidOperationException("Hedef klasör boş olmalı.");

        progress?.Report("Proje kopyalanıyor...");
        CopyDirectory(source, target);

        foreach (var fr in result.Files.Where(f => f.HasChanges && f.File.IsInsideRoot))
        {
            progress?.Report($"Yazılıyor: {fr.File.RelativePath}");
            var lines = ApplyEdits(fr);
            var sb = new StringBuilder();
            for (int i = 0; i < lines.Count; i++)
            {
                sb.Append(lines[i]);
                if (i < lines.Count - 1 || fr.File.EndsWithNewLine) sb.Append(fr.File.NewLine);
            }
            File.WriteAllText(Path.Combine(target, fr.File.RelativePath!), sb.ToString(), VbProject.FileEncoding);
        }

        File.WriteAllText(Path.Combine(target, ReportFileName), CleanupReport.Build(result), Encoding.UTF8);
        progress?.Report("Tamamlandı.");
    }

    static void CopyDirectory(string source, string target)
    {
        Directory.CreateDirectory(target);
        foreach (var dir in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(Path.Combine(target, Path.GetRelativePath(source, dir)));
        foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
            File.Copy(file, Path.Combine(target, Path.GetRelativePath(source, file)));
    }

    public static List<string> ApplyEdits(FileResult fr)
    {
        var f = fr.File;
        var lines = f.Lines;
        var delete = new bool[lines.Count];
        var replace = new Dictionary<int, (int Last, string Text)>();

        foreach (var node in fr.RemovalRoots)
            for (int i = node.BeginLine; i <= node.EndLine; i++) delete[i] = true;

        foreach (var proc in fr.RemovedProcs.Where(p => p.EndLine >= 0))
        {
            for (int i = proc.HeaderLine; i <= proc.EndLine; i++) delete[i] = true;
            // Avoid leaving two blank lines where the procedure was.
            int after = proc.EndLine + 1, before = proc.HeaderLine - 1;
            if (after < lines.Count && before >= 0 && IsBlank(lines[after]) && IsBlank(lines[before])) delete[after] = true;
        }

        var deadUnits = fr.RemovedUnits.ToHashSet();
        foreach (var line in deadUnits.Select(u => u.Line).Distinct())
        {
            if (delete[line.FirstLine]) continue;
            var live = line.Units.Where(u => !deadUnits.Contains(u)).ToList();
            string? text;
            if (live.Count == 0)
                text = line.Label != null ? line.Indent + line.Label + ":" + Suffix(line.Comment) : null;
            else
                text = line.Indent + (line.Label != null ? line.Label + ": " : "")
                       + string.Join(": ", live.Select(u => u.Text)) + Suffix(line.Comment);

            if (text == null)
                for (int i = line.FirstLine; i <= line.LastLine; i++) delete[i] = true;
            else
                replace[line.FirstLine] = (line.LastLine, text);
        }

        var output = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++)
        {
            if (replace.TryGetValue(i, out var r))
            {
                output.Add(r.Text);
                i = r.Last;
                continue;
            }
            if (!delete[i]) output.Add(lines[i]);
        }
        return output;

        static string Suffix(string comment) => comment.Length > 0 ? " " + comment : "";
        static bool IsBlank(string s) => s.Trim().Length == 0;
    }
}

public static class CleanupReport
{
    public static string Build(AnalysisResult r)
    {
        var sb = new StringBuilder();
        sb.AppendLine("VB Screen Cleaner raporu");
        sb.AppendLine($"Kaynak: {r.Project.RootDir}");
        sb.AppendLine($"Tarih: {DateTime.Now:yyyy-MM-dd HH:mm}");
        sb.AppendLine($"Form_Load içinde gizlenenler dahil: {(r.Options.IncludeHiddenInFormLoad ? "Evet" : "Hayır")}");
        sb.AppendLine($"Controls koleksiyonu kullanan formlar atlandı: {(r.Options.SkipFormsUsingControlsCollection ? "Evet" : "Hayır")}");
        sb.AppendLine();
        sb.AppendLine($"Silinen kontrol sayısı: {r.TotalRemovedControls}");
        sb.AppendLine($"Silinen olay prosedürü: {r.Files.Sum(f => f.RemovedProcs.Count)}");
        sb.AppendLine($"Silinen kod satırı/ifadesi: {r.Files.Sum(f => f.RemovedUnits.Count)}");
        sb.AppendLine($"Değişen dosya: {r.Files.Count(f => f.HasChanges)}");
        sb.AppendLine();

        if (r.Warnings.Count > 0)
        {
            sb.AppendLine("UYARILAR");
            foreach (var w in r.Warnings) sb.AppendLine("  ! " + w);
            sb.AppendLine();
        }

        foreach (var fr in r.Files.Where(f => f.HasChanges || f.KeptHidden.Count > 0 || f.Notes.Count > 0))
        {
            sb.AppendLine(new string('=', 78));
            sb.AppendLine($"{fr.File.DisplayName}  ({fr.File.ObjectName})");
            foreach (var n in fr.Notes) sb.AppendLine("  Not: " + n);

            if (fr.RemovalRoots.Count > 0)
            {
                sb.AppendLine("  SİLİNEN KONTROLLER:");
                foreach (var n in fr.RemovalRoots)
                {
                    sb.AppendLine($"    - {n.DisplayName} [{n.TypeName}] satır {n.BeginLine + 1}: {n.HiddenWhy}");
                    foreach (var d in n.Descendants())
                        sb.AppendLine($"        · {d.DisplayName} [{d.TypeName}]");
                }
            }
            if (fr.RemovedProcs.Count > 0)
            {
                sb.AppendLine("  SİLİNEN OLAY PROSEDÜRLERİ:");
                foreach (var p in fr.RemovedProcs) sb.AppendLine($"    - {p.Name} (satır {p.HeaderLine + 1})");
            }
            if (fr.RemovedUnits.Count > 0)
            {
                sb.AppendLine("  SİLİNEN KOD İFADELERİ:");
                foreach (var u in fr.RemovedUnits) sb.AppendLine($"    - satır {u.Line.FirstLine + 1}: {u.Text}");
            }
            if (fr.KeptHidden.Count > 0)
            {
                sb.AppendLine("  GİZLİ AMA KORUNAN KONTROLLER:");
                foreach (var n in fr.KeptHidden)
                {
                    sb.AppendLine($"    - {n.DisplayName} [{n.TypeName}]: {n.HiddenWhy}");
                    foreach (var reason in CleanupAnalyzer.KeepReasons(n).Take(5)) sb.AppendLine($"        neden: {reason}");
                }
            }
            sb.AppendLine();
        }
        return sb.ToString();
    }
}
