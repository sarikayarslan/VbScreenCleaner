using System.Text;

namespace VbOrphanHandlers;

public sealed class MainForm : Form
{
    readonly TextBox txtProject = new() { Dock = DockStyle.Fill };
    readonly TextBox txtLib = new() { Dock = DockStyle.Fill, PlaceholderText = "İsteğe bağlı; birden fazla klasör ; ile" };
    readonly TextBox txtReport = new()
    {
        Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
        WordWrap = false, Font = new Font("Consolas", 9.5f),
    };
    readonly Button btnScan = new() { Text = "Tara", AutoSize = true };
    readonly Button btnSave = new() { Text = "Raporu kaydet...", AutoSize = true, Enabled = false };
    readonly Label lblStatus = new() { AutoSize = true, Anchor = AnchorStyles.Left };

    public MainForm()
    {
        Text = "VB5 Sahipsiz Olay Prosedürleri";
        Width = 1000;
        Height = 700;
        StartPosition = FormStartPosition.CenterScreen;

        var btnFolder = new Button { Text = "Klasör...", AutoSize = true };
        var btnVbp = new Button { Text = ".vbp...", AutoSize = true };
        var btnLib = new Button { Text = "Ekle...", AutoSize = true };
        btnFolder.Click += (_, _) => { var p = PickFolder(); if (p != null) txtProject.Text = p; };
        btnVbp.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "VB projesi (*.vbp)|*.vbp" };
            if (dlg.ShowDialog(this) == DialogResult.OK) txtProject.Text = dlg.FileName;
        };
        btnLib.Click += (_, _) =>
        {
            var p = PickFolder();
            if (p != null) txtLib.Text = string.IsNullOrWhiteSpace(txtLib.Text) ? p : txtLib.Text.TrimEnd(';', ' ') + ";" + p;
        };
        btnScan.Click += async (_, _) => await ScanAsync();
        btnSave.Click += (_, _) => SaveReport();

        var grid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Padding = new Padding(8) };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.Controls.Add(new Label { Text = "Proje (klasör / .vbp):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        grid.Controls.Add(txtProject, 1, 0);
        grid.Controls.Add(btnFolder, 2, 0);
        grid.Controls.Add(btnVbp, 3, 0);
        grid.Controls.Add(new Label { Text = "Lib klasörü:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        grid.Controls.Add(txtLib, 1, 1);
        grid.Controls.Add(btnLib, 2, 1);

        var actions = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8, 0, 8, 4) };
        actions.Controls.AddRange([btnScan, btnSave, lblStatus]);

        var reportPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };
        reportPanel.Controls.Add(txtReport);

        Controls.Add(reportPanel);
        Controls.Add(actions);
        Controls.Add(grid);
        AcceptButton = btnScan;
    }

    string? PickFolder()
    {
        using var dlg = new FolderBrowserDialog();
        return dlg.ShowDialog(this) == DialogResult.OK ? dlg.SelectedPath : null;
    }

    async Task ScanAsync()
    {
        var inputs = new List<string>();
        if (!string.IsNullOrWhiteSpace(txtProject.Text)) inputs.Add(txtProject.Text.Trim().Trim('"'));
        inputs.AddRange(txtLib.Text.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => s.Trim('"')));
        if (inputs.Count == 0)
        {
            MessageBox.Show(this, "Proje klasörünü ya da .vbp dosyasını seçin.", Text);
            return;
        }

        btnScan.Enabled = btnSave.Enabled = false;
        lblStatus.Text = "Taranıyor...";
        try
        {
            txtReport.Text = (await Task.Run(() => OrphanScanner.Scan(inputs))).Replace("\n", "\r\n").Replace("\r\r\n", "\r\n");
            lblStatus.Text = "Tamamlandı.";
            btnSave.Enabled = true;
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Hata.";
            MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnScan.Enabled = true;
        }
    }

    void SaveReport()
    {
        using var dlg = new SaveFileDialog { Filter = "Metin dosyası (*.txt)|*.txt", FileName = "orphan-rapor.txt" };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            File.WriteAllText(dlg.FileName, txtReport.Text, new UTF8Encoding(true));
    }
}
