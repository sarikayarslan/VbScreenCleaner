using System.Diagnostics;
using VbScreenCleaner.Core;

namespace VbScreenCleaner;

public sealed class MainForm : Form
{
    readonly TextBox _source = new() { Dock = DockStyle.Fill };
    readonly TextBox _target = new() { Dock = DockStyle.Fill };
    readonly CheckBox _loadHidden = new() { Text = "Form_Load içinde koşulsuz gizlenen kontrolleri de temizle", Checked = true, AutoSize = true };
    readonly CheckBox _skipControls = new() { Text = "Controls koleksiyonu kullanan formlara dokunma", Checked = true, AutoSize = true };
    readonly CheckBox _deadCode = new() { Text = "Hiç çağrılmayan prosedürleri ölü kod say (gerekirse sil)", Checked = true, AutoSize = true };
    readonly Button _analyze = new() { Text = "Analiz Et", AutoSize = true };
    readonly Button _clean = new() { Text = "Temizle ve Kaydet", AutoSize = true, Enabled = false };
    readonly Button _checkAll = new() { Text = "Tümünü Seç", AutoSize = true, Enabled = false };
    readonly Button _uncheckAll = new() { Text = "Hiçbirini Seçme", AutoSize = true, Enabled = false };
    readonly TreeView _tree = new() { Dock = DockStyle.Fill, CheckBoxes = true, HideSelection = false };
    readonly TextBox _details = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Font = new Font("Consolas", 9.5f) };
    readonly ToolStripStatusLabel _status = new() { Text = "Kaynak proje klasörünü seçin." };

    AnalysisResult? _result;
    bool _syncingChecks;

    public MainForm()
    {
        Text = "VB Screen Cleaner — VB5 gizli ekran öğesi temizleyici";
        Width = 1200;
        Height = 800;
        StartPosition = FormStartPosition.CenterScreen;

        var paths = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3, Padding = new Padding(8) };
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        AddPathRow(paths, 0, "Kaynak proje klasörü:", _source);
        AddPathRow(paths, 1, "Hedef klasör (boş):", _target);

        var options = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8, 0, 8, 4) };
        options.Controls.AddRange(new Control[] { _loadHidden, _skipControls, _deadCode });

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8, 0, 8, 8) };
        buttons.Controls.AddRange(new Control[] { _analyze, _clean, _checkAll, _uncheckAll });

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 520 };
        split.Panel1.Controls.Add(_tree);
        split.Panel2.Controls.Add(_details);

        var statusStrip = new StatusStrip();
        statusStrip.Items.Add(_status);

        Controls.Add(split);
        Controls.Add(buttons);
        Controls.Add(options);
        Controls.Add(paths);
        Controls.Add(statusStrip);

        _analyze.Click += async (_, _) => await AnalyzeAsync();
        _clean.Click += async (_, _) => await CleanAsync();
        _checkAll.Click += (_, _) => SetAllChecks(true);
        _uncheckAll.Click += (_, _) => SetAllChecks(false);
        _tree.BeforeCheck += (_, e) => { if (!_syncingChecks && e.Node?.Tag is not ControlNode && e.Node?.Tag is not FileResult) e.Cancel = true; };
        _tree.AfterCheck += OnAfterCheck;
        _tree.AfterSelect += (_, e) => ShowDetails(e.Node);
        _loadHidden.CheckedChanged += (_, _) => InvalidateAnalysis();
        _skipControls.CheckedChanged += (_, _) => InvalidateAnalysis();
        _deadCode.CheckedChanged += (_, _) => InvalidateAnalysis();
        _source.TextChanged += (_, _) => InvalidateAnalysis();
    }

    void AddPathRow(TableLayoutPanel panel, int row, string label, TextBox box)
    {
        var browse = new Button { Text = "Seç...", AutoSize = true };
        browse.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog { SelectedPath = box.Text, UseDescriptionForTitle = true, Description = label };
            if (dlg.ShowDialog(this) == DialogResult.OK) box.Text = dlg.SelectedPath;
        };
        panel.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 7, 3, 3) }, 0, row);
        panel.Controls.Add(box, 1, row);
        panel.Controls.Add(browse, 2, row);
    }

    CleanupOptions BuildOptions(IEnumerable<string>? forcedKeep = null)
    {
        var o = new CleanupOptions
        {
            IncludeHiddenInFormLoad = _loadHidden.Checked,
            SkipFormsUsingControlsCollection = _skipControls.Checked,
            RemoveDeadCode = _deadCode.Checked,
        };
        if (forcedKeep != null) o.ForcedKeep.UnionWith(forcedKeep);
        return o;
    }

    void InvalidateAnalysis()
    {
        if (_result == null) return;
        _result = null;
        _tree.Nodes.Clear();
        _details.Clear();
        _clean.Enabled = _checkAll.Enabled = _uncheckAll.Enabled = false;
        _status.Text = "Ayarlar değişti, yeniden analiz edin.";
    }

    async Task AnalyzeAsync()
    {
        var src = _source.Text.Trim();
        if (!Directory.Exists(src))
        {
            MessageBox.Show(this, "Geçerli bir kaynak proje klasörü seçin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(_target.Text))
            _target.Text = Path.GetFullPath(src).TrimEnd('\\') + "_Temiz";

        SetBusy(true, "Analiz ediliyor...");
        try
        {
            var options = BuildOptions();
            _result = await Task.Run(() => CleanupAnalyzer.Analyze(VbProject.Load(src), options));
            FillTree(_result);
            _status.Text = $"{_result.Project.Files.Count(f => f.HasDesigner)} ekran tarandı · {_result.TotalRemovedControls} kontrol silinecek · "
                           + $"{_result.Files.Sum(f => f.KeptHidden.Count)} gizli kontrol korunuyor · {_result.Warnings.Count} uyarı";
            _details.Text = CleanupReport.Build(_result).Replace("\n", "\r\n");
        }
        catch (Exception ex)
        {
            _result = null;
            MessageBox.Show(this, ex.Message, "Analiz hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _status.Text = "Analiz başarısız.";
        }
        finally
        {
            SetBusy(false, null);
            _clean.Enabled = _checkAll.Enabled = _uncheckAll.Enabled = _result != null;
        }
    }

    async Task CleanAsync()
    {
        if (_result == null) return;
        var target = _target.Text.Trim();
        if (target.Length == 0)
        {
            MessageBox.Show(this, "Hedef klasörü seçin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Unchecked removal items are kept; re-analyze so dependent decisions follow.
        var keep = AllNodes(_tree.Nodes)
            .Where(n => n.Tag is ControlNode { IsRemovalRoot: true } && !n.Checked)
            .Select(n => ((ControlNode)n.Tag!).Group!.Key)
            .ToList();

        var src = _result.Project.RootDir;
        SetBusy(true, "Temizleniyor...");
        try
        {
            var options = BuildOptions(keep);
            var progress = new Progress<string>(s => _status.Text = s);
            var final = await Task.Run(() =>
            {
                var r = CleanupAnalyzer.Analyze(VbProject.Load(src), options);
                CleanupWriter.Write(r, target, progress);
                return r;
            });

            var outside = final.Project.Files.Count(f => !f.IsInsideRoot);
            var msg = $"Temizlik tamamlandı.\n\nSilinen kontrol: {final.TotalRemovedControls}\n"
                      + $"Silinen olay prosedürü: {final.Files.Sum(f => f.RemovedProcs.Count)}\n"
                      + $"Silinen kullanılmayan prosedür: {final.Files.Sum(f => f.DeadCodeProcs.Count)}\n"
                      + $"Silinen kod ifadesi: {final.Files.Sum(f => f.RemovedUnits.Count)}\n\n"
                      + $"Rapor: {CleanupWriter.ReportFileName}"
                      + (outside > 0 ? $"\n\nDikkat: {outside} dosya proje klasörü dışında olduğu için kopyalanmadı (raporda listelendi)." : "")
                      + "\n\nHedef klasör açılsın mı?";
            if (MessageBox.Show(this, msg, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{target}\"") { UseShellExecute = true });
            _status.Text = $"Kaydedildi: {target}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Temizleme hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _status.Text = "Temizleme başarısız.";
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    void SetBusy(bool busy, string? status)
    {
        UseWaitCursor = busy;
        _analyze.Enabled = !busy;
        _clean.Enabled = !busy && _result != null;
        if (status != null) _status.Text = status;
    }

    void FillTree(AnalysisResult r)
    {
        _syncingChecks = true;
        _tree.BeginUpdate();
        _tree.Nodes.Clear();

        if (r.Warnings.Count > 0)
        {
            var w = _tree.Nodes.Add($"Uyarılar ({r.Warnings.Count})");
            w.ForeColor = Color.DarkOrange;
            foreach (var text in r.Warnings) w.Nodes.Add(text).ForeColor = Color.DarkOrange;
        }

        foreach (var fr in r.Files.Where(f => f.File.HasDesigner && (f.HasChanges || f.KeptHidden.Count > 0 || f.Notes.Count > 0))
                                  .OrderBy(f => f.File.DisplayName))
        {
            var fileNode = _tree.Nodes.Add($"{fr.File.DisplayName}  —  {fr.RemovedControlCount} silinecek, {fr.KeptHidden.Count} korunan gizli");
            fileNode.Tag = fr;
            fileNode.Checked = fr.RemovalRoots.Count > 0;

            foreach (var n in fr.RemovalRoots)
            {
                int sub = n.Descendants().Count();
                var tn = fileNode.Nodes.Add($"{n.DisplayName}  [{n.TypeName}]  {n.HiddenWhy}{(sub > 0 ? $"  (+{sub} alt kontrol)" : "")}");
                tn.Tag = n;
                tn.Checked = true;
            }
            foreach (var n in fr.KeptHidden)
            {
                var tn = fileNode.Nodes.Add($"KORUNUYOR: {n.DisplayName}  [{n.TypeName}]  {CleanupAnalyzer.KeepReasons(n)[0]}");
                tn.Tag = new KeptInfo(n);
                tn.ForeColor = Color.Gray;
            }
            foreach (var note in fr.Notes) fileNode.Nodes.Add("Not: " + note).ForeColor = Color.DarkOrange;
        }

        _tree.EndUpdate();
        _syncingChecks = false;
    }

    void OnAfterCheck(object? sender, TreeViewEventArgs e)
    {
        if (_syncingChecks || e.Node == null) return;
        _syncingChecks = true;
        if (e.Node.Tag is FileResult)
        {
            foreach (TreeNode child in e.Node.Nodes)
                if (child.Tag is ControlNode) child.Checked = e.Node.Checked;
        }
        else if (e.Node.Tag is ControlNode && e.Node.Parent != null)
        {
            e.Node.Parent.Checked = e.Node.Parent.Nodes.Cast<TreeNode>().Any(n => n.Tag is ControlNode && n.Checked);
        }
        _syncingChecks = false;
    }

    void SetAllChecks(bool value)
    {
        _syncingChecks = true;
        foreach (var n in AllNodes(_tree.Nodes).Where(n => n.Tag is ControlNode || n.Tag is FileResult)) n.Checked = value;
        _syncingChecks = false;
    }

    void ShowDetails(TreeNode? node)
    {
        switch (node?.Tag)
        {
            case FileResult fr:
                _details.Text = FileDetails(fr);
                break;
            case ControlNode n:
                _details.Text = NodeDetails(n, kept: false);
                break;
            case KeptInfo k:
                _details.Text = NodeDetails(k.Node, kept: true);
                break;
        }
    }

    static string FileDetails(FileResult fr)
    {
        var lines = new List<string> { fr.File.FullPath, "" };
        lines.AddRange(fr.Notes.Select(n => "Not: " + n));
        lines.Add($"Silinecek kontrol: {fr.RemovedControlCount}");
        lines.Add($"Silinecek olay prosedürü: {fr.RemovedProcs.Count}");
        foreach (var p in fr.RemovedProcs) lines.Add($"   {p.Name}  (satır {p.HeaderLine + 1})");
        lines.Add($"Silinecek kullanılmayan prosedür: {fr.DeadCodeProcs.Count}");
        foreach (var p in fr.DeadCodeProcs) lines.Add($"   {p.Name}  (satır {p.HeaderLine + 1})");
        foreach (var (owner, prop) in fr.TabListRemovals) lines.Add($"Sekme listesi güncellenecek: {owner.DisplayName}.{prop}");
        lines.Add($"Silinecek kod ifadesi: {fr.RemovedUnits.Count}");
        foreach (var u in fr.RemovedUnits) lines.Add($"   satır {u.Line.FirstLine + 1}: {u.Text}");
        return string.Join("\r\n", lines);
    }

    static string NodeDetails(ControlNode n, bool kept)
    {
        var f = n.Group!.File;
        var lines = new List<string>
        {
            $"{f.ObjectName}.{n.DisplayName}   [{n.TypeName}]",
            $"Dosya: {f.DisplayName}, satır {n.BeginLine + 1}-{n.EndLine + 1}",
            $"Gizli olma nedeni: {n.HiddenWhy}",
            "",
        };
        if (kept)
        {
            lines.Add("Neden korunuyor:");
            lines.AddRange(CleanupAnalyzer.KeepReasons(n).Select(r => "   • " + r));
        }
        else
        {
            var sub = n.Descendants().ToList();
            if (sub.Count > 0)
            {
                lines.Add("Birlikte silinecek alt kontroller:");
                lines.AddRange(sub.Select(d => $"   {d.DisplayName}  [{d.TypeName}]"));
                lines.Add("");
            }
            var groups = n.SelfAndDescendants().Select(d => d.Group!).ToHashSet();
            var procs = groups.SelectMany(g => g.Handlers).ToList();
            if (procs.Count > 0)
            {
                lines.Add("Silinecek olay prosedürleri:");
                lines.AddRange(procs.Select(p => $"   {p.Name}  (satır {p.HeaderLine + 1})"));
                lines.Add("");
            }
            lines.Add("İşaretini kaldırırsanız bu kontrol (ve içindekiler) korunur.");
        }
        lines.Add("");
        lines.Add("Tasarım özellikleri:");
        lines.AddRange(n.Properties.Select(p => $"   {p.Key} = {p.Value.Trim()}"));
        return string.Join("\r\n", lines);
    }

    static IEnumerable<TreeNode> AllNodes(TreeNodeCollection nodes)
    {
        foreach (TreeNode n in nodes)
        {
            yield return n;
            foreach (var c in AllNodes(n.Nodes)) yield return c;
        }
    }

    sealed record KeptInfo(ControlNode Node);
}
