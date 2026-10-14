using ExcelPdfConverter.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
namespace ExcelPdfConverter
{
    public sealed class MainForm : Form
    {
        private readonly ListView files = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, AllowDrop = true };
        private readonly TextBox output = new TextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly Button add = new Button { Text = "Excelを追加", AutoSize = true };
        private readonly Button remove = new Button { Text = "選択を削除", AutoSize = true };
        private readonly Button browse = new Button { Text = "保存先を選択", AutoSize = true };
        private readonly Button start = new Button { Text = "PDFに変換", AutoSize = true };
        private readonly Button cancel = new Button { Text = "中止", AutoSize = true, Enabled = false };
        private readonly Button open = new Button { Text = "保存先を開く", AutoSize = true };
        private readonly CheckBox ignoreArea = new CheckBox { Text = "印刷範囲を無視する", AutoSize = true };
        private readonly Label summary = new Label { Text = "Excelファイルを追加してください（ドラッグ＆ドロップ対応）", AutoSize = true };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Fill };
        private bool running;
        private volatile bool stopRequested;
        public MainForm()
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            Text = "Excel → PDF 変換"; ClientSize = new Size(860, 570); MinimumSize = new Size(680, 460);
            Font = new Font("Yu Gothic UI", 10); StartPosition = FormStartPosition.CenterScreen;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 1, RowCount = 7 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            layout.Controls.Add(new Label { Text = "Excelを、印刷レイアウトのままPDFへ\n複数ファイルを一括変換 · 元ファイルと既存PDFは上書きしません", AutoSize = true }, 0, 0);
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            toolbar.Controls.AddRange(new Control[] { add, remove, ignoreArea }); layout.Controls.Add(toolbar, 0, 1);
            files.Columns.Add("Excelファイル", 230); files.Columns.Add("場所 / 出力先", 375); files.Columns.Add("結果", 160);
            layout.Controls.Add(files, 0, 2);
            var destination = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            destination.Controls.Add(output, 0, 0); destination.Controls.Add(browse, 1, 0); layout.Controls.Add(destination, 0, 3);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill }; actions.Controls.AddRange(new Control[] { start, cancel, open }); layout.Controls.Add(actions, 0, 4);
            layout.Controls.Add(progress, 0, 5); layout.Controls.Add(summary, 0, 6); Controls.Add(layout);
            output.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Excel PDFs");
            add.Click += (s, e) => { using (var dialog = new OpenFileDialog { Filter = "Excelブック|*.xlsx;*.xls;*.xlsm;*.xlsb", Multiselect = true }) if (dialog.ShowDialog(this) == DialogResult.OK) AddFiles(dialog.FileNames); };
            remove.Click += (s, e) => { foreach (ListViewItem item in files.SelectedItems) files.Items.Remove(item); };
            browse.Click += (s, e) => { using (var dialog = new FolderBrowserDialog { Description = "PDFの保存先", SelectedPath = output.Text }) if (dialog.ShowDialog(this) == DialogResult.OK) output.Text = dialog.SelectedPath; };
            files.DragEnter += (s, e) => e.Effect = !running && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            files.DragDrop += (s, e) => { if (!running) AddFiles((string[])e.Data.GetData(DataFormats.FileDrop)); };
            start.Click += (s, e) => StartConversion();
            cancel.Click += (s, e) => { stopRequested = true; cancel.Enabled = false; summary.Text = "現在のファイルの変換終了後に中止します…"; };
            open.Click += (s, e) => { if (Directory.Exists(output.Text)) Process.Start(new ProcessStartInfo { FileName = output.Text, UseShellExecute = true }); else MessageBox.Show(this, "保存先はまだ作成されていません。"); };
            FormClosing += (s, e) => { if (running) { e.Cancel = true; MessageBox.Show(this, "変換中です。「中止」を押し、処理の終了を待ってください。"); } };
        }
        private void AddFiles(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                if (!File.Exists(path) || !OutputPaths.IsSupported(path) || Path.GetFileName(path).StartsWith("~$")) continue;
                string full = Path.GetFullPath(path); bool duplicate = false;
                foreach (ListViewItem existing in files.Items) if (string.Equals((string)existing.Tag, full, StringComparison.OrdinalIgnoreCase)) duplicate = true;
                if (!duplicate) files.Items.Add(new ListViewItem(new[] { Path.GetFileName(full), full, "待機" }) { Tag = full });
            }
            summary.Text = files.Items.Count + " 件のファイル";
        }
        private void StartConversion()
        {
            if (running || files.Items.Count == 0) { MessageBox.Show(this, "Excelファイルを追加してください。"); return; }
            try { Directory.CreateDirectory(output.Text); } catch (Exception ex) { MessageBox.Show(this, "保存先を作成できません。\n" + ex.Message); return; }
            var items = new List<ListViewItem>(); foreach (ListViewItem item in files.Items) items.Add(item);
            string directory = output.Text; bool ignore = ignoreArea.Checked;
            running = true; stopRequested = false;
            add.Enabled = remove.Enabled = browse.Enabled = start.Enabled = ignoreArea.Enabled = false; cancel.Enabled = true;
            progress.Value = 0; progress.Maximum = items.Count;
            foreach (var item in items) item.SubItems[2].Text = "待機";
            var thread = new Thread(() =>
            {
                int success = 0, failed = 0, completed = 0;
                foreach (var item in items)
                {
                    if (stopRequested) break;
                    UpdateUi(() => { item.SubItems[2].Text = "変換中"; summary.Text = "変換中: " + item.Text; });
                    try
                    {
                        string result = PdfConverter.Convert((string)item.Tag, directory, ignore);
                        success++;
                        UpdateUi(() => { item.SubItems[1].Text = result; item.SubItems[2].Text = "完了"; item.ForeColor = Color.DarkGreen; item.ToolTipText = result; });
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        UpdateUi(() => { item.SubItems[2].Text = "失敗"; item.ForeColor = Color.Firebrick; item.SubItems[1].Text = ex.Message; item.ToolTipText = ex.ToString(); });
                    }
                    completed++; int value = completed; UpdateUi(() => progress.Value = value);
                }
                UpdateUi(() =>
                {
                    running = false; add.Enabled = remove.Enabled = browse.Enabled = start.Enabled = ignoreArea.Enabled = true; cancel.Enabled = false;
                    summary.Text = (stopRequested ? "中止しました" : "処理が終了しました") + "：成功 " + success + " / 失敗 " + failed + " / 未処理 " + (items.Count - completed);
                });
            }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
        }
        private void UpdateUi(Action action) { Invoke(action); }
    }
}
