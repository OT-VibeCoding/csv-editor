// CSVエディタ - シンプル・軽量な CSV 編集ソフト (Windows 11 / .NET Framework 4.x 標準搭載)
// ビルド: build.bat を実行 (追加インストール不要)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("CSVエディタ")]
[assembly: AssemblyProduct("CSVエディタ")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace CsvEditor
{
    static class Program
    {
        [DllImport("user32.dll")]
        static extern bool SetProcessDPIAware();

        // 予期しないエラーは内容を表示する (Ctrl+C でメッセージをコピー可能)
        static void ShowError(Exception ex)
        {
            string msg = ex == null ? "不明なエラー" : ex.ToString();
            try
            {
                string log = Path.Combine(Path.GetTempPath(), "CsvEditor_error.txt");
                File.WriteAllText(log, DateTime.Now + "\r\n" + msg);
                msg += "\r\n\r\n(このエラー内容は " + log + " にも保存しました)";
            }
            catch { }
            MessageBox.Show("エラーが発生しました。\r\n\r\n" + msg, "CSVエディタ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        [STAThread]
        static void Main(string[] args)
        {
            try { if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) => ShowError(e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => ShowError(e.ExceptionObject as Exception);
            MainForm f;
            try { f = new MainForm(); }
            catch (Exception ex) { ShowError(ex); return; }
            if (args.Length > 0) f.PendingOpen = args[0];
            Application.Run(f);
        }
    }

    // ===================== データモデル =====================
    class Row
    {
        public List<string> V;
        public Row(int n) { V = new List<string>(n); for (int i = 0; i < n; i++) V.Add(""); }
        public Row(List<string> v) { V = v; }
    }

    class Sheet
    {
        public List<Row> Rows = new List<Row>();
        public int Cols;

        public string Get(int r, int c)
        {
            if (r >= 0 && c >= 0 && r < Rows.Count && c < Cols) return Rows[r].V[c];
            return "";
        }
        public void Set(int r, int c, string v)
        {
            EnsureSize(r + 1, c + 1);
            Rows[r].V[c] = v ?? "";
        }
        public void EnsureSize(int rows, int cols)
        {
            if (cols > Cols)
            {
                foreach (var row in Rows)
                {
                    while (row.V.Count < cols) row.V.Add("");
                }
                Cols = cols;
            }
            while (Rows.Count < rows) Rows.Add(new Row(Cols));
        }
        public void Truncate(int rows, int cols)
        {
            if (Rows.Count > rows) Rows.RemoveRange(rows, Rows.Count - rows);
            if (cols < Cols)
            {
                foreach (var row in Rows) row.V.RemoveRange(cols, row.V.Count - cols);
                Cols = cols;
            }
        }
        public void Normalize()
        {
            int m = 0;
            foreach (var r in Rows) if (r.V.Count > m) m = r.V.Count;
            Cols = m;
            foreach (var r in Rows) while (r.V.Count < m) r.V.Add("");
        }
    }

    // ===================== CSV 入出力 =====================
    enum EncKind { Utf8Bom, Utf8, ShiftJis, Utf16 }

    static class Csv
    {
        public static List<List<string>> Parse(string text, char delim)
        {
            var rows = new List<List<string>>();
            var row = new List<string>();
            var sb = new StringBuilder();
            bool inQ = false, quoted = false;
            int i = 0, n = text.Length;
            while (i < n)
            {
                char ch = text[i];
                if (inQ)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < n && text[i + 1] == '"') { sb.Append('"'); i += 2; continue; }
                        inQ = false; i++; continue;
                    }
                    sb.Append(ch); i++; continue;
                }
                if (ch == '"' && sb.Length == 0 && !quoted) { inQ = true; quoted = true; i++; continue; }
                if (ch == delim) { row.Add(sb.ToString()); sb.Length = 0; quoted = false; i++; continue; }
                if (ch == '\r' || ch == '\n')
                {
                    row.Add(sb.ToString()); sb.Length = 0; quoted = false;
                    rows.Add(row); row = new List<string>();
                    if (ch == '\r' && i + 1 < n && text[i + 1] == '\n') i++;
                    i++; continue;
                }
                sb.Append(ch); i++;
            }
            if (sb.Length > 0 || quoted || row.Count > 0) { row.Add(sb.ToString()); rows.Add(row); }
            return rows;
        }

        public static string Escape(string s, char delim)
        {
            if (s.IndexOf(delim) >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\r') >= 0 || s.IndexOf('\n') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        public static char DetectDelimiter(string text, string path)
        {
            string ext = (Path.GetExtension(path ?? "") ?? "").ToLowerInvariant();
            if (ext == ".tsv" || ext == ".tab") return '\t';
            int comma = 0, tab = 0, semi = 0; bool q = false;
            int lines = 0;
            for (int i = 0; i < text.Length && i < 65536; i++)
            {
                char c = text[i];
                if (c == '"') q = !q;
                else if (!q)
                {
                    if (c == ',') comma++;
                    else if (c == '\t') tab++;
                    else if (c == ';') semi++;
                    else if (c == '\n' && ++lines >= 20) break;
                }
            }
            if (tab > comma && tab >= semi) return '\t';
            if (semi > comma && semi > tab) return ';';
            return ',';
        }

        public static string EncName(EncKind k)
        {
            switch (k)
            {
                case EncKind.Utf8Bom: return "UTF-8 (BOM付き)";
                case EncKind.Utf8: return "UTF-8";
                case EncKind.ShiftJis: return "Shift_JIS";
                default: return "UTF-16 LE";
            }
        }

        public static string Decode(byte[] b, EncKind? forced, out EncKind kind)
        {
            if (forced.HasValue)
            {
                kind = forced.Value;
                int skip = 0;
                if (kind == EncKind.Utf8 || kind == EncKind.Utf8Bom)
                {
                    if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { skip = 3; kind = EncKind.Utf8Bom; }
                    return new UTF8Encoding(false).GetString(b, skip, b.Length - skip);
                }
                if (kind == EncKind.Utf16)
                {
                    if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) skip = 2;
                    return Encoding.Unicode.GetString(b, skip, b.Length - skip);
                }
                return Encoding.GetEncoding(932).GetString(b);
            }
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF)
            { kind = EncKind.Utf8Bom; return new UTF8Encoding(false).GetString(b, 3, b.Length - 3); }
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE)
            { kind = EncKind.Utf16; return Encoding.Unicode.GetString(b, 2, b.Length - 2); }
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF)
            { kind = EncKind.Utf16; return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2); }
            try
            {
                var s = new UTF8Encoding(false, true).GetString(b);
                kind = EncKind.Utf8; return s;
            }
            catch (DecoderFallbackException) { }
            kind = EncKind.ShiftJis;
            return Encoding.GetEncoding(932).GetString(b);
        }

        // strict=true のとき Shift_JIS で表せない文字があれば EncoderFallbackException
        public static byte[] Encode(string text, EncKind k, bool strict)
        {
            byte[] pre = new byte[0];
            Encoding e;
            switch (k)
            {
                case EncKind.Utf8Bom: e = new UTF8Encoding(false); pre = new byte[] { 0xEF, 0xBB, 0xBF }; break;
                case EncKind.Utf8: e = new UTF8Encoding(false); break;
                case EncKind.Utf16: e = new UnicodeEncoding(false, false); pre = new byte[] { 0xFF, 0xFE }; break;
                default:
                    e = strict ? Encoding.GetEncoding(932, EncoderFallback.ExceptionFallback, DecoderFallback.ReplacementFallback)
                               : Encoding.GetEncoding(932);
                    break;
            }
            byte[] body = e.GetBytes(text);
            if (pre.Length == 0) return body;
            var all = new byte[pre.Length + body.Length];
            Buffer.BlockCopy(pre, 0, all, 0, pre.Length);
            Buffer.BlockCopy(body, 0, all, pre.Length, body.Length);
            return all;
        }
    }

    // ===================== 元に戻す =====================
    class UndoItem
    {
        public string Name;
        public Action Undo, Redo;
    }

    class CellChange { public int R, C; public string Old, New; }

    // 並べ替え用比較 (数値は数値として、空欄は常に末尾)
    class SmartComparer : IComparer<string>
    {
        readonly CompareInfo ci = CultureInfo.CurrentCulture.CompareInfo;
        public int Compare(string a, string b)
        {
            double x, y;
            bool nx = double.TryParse(a, NumberStyles.Any, CultureInfo.CurrentCulture, out x);
            bool ny = double.TryParse(b, NumberStyles.Any, CultureInfo.CurrentCulture, out y);
            if (nx && ny) return x.CompareTo(y);
            if (nx) return -1;
            if (ny) return 1;
            return ci.Compare(a, b, CompareOptions.IgnoreCase | CompareOptions.IgnoreWidth);
        }
    }

    // ===================== メインウィンドウ =====================
    class MainForm : Form
    {
        const string AppName = "CSVエディタ";
        const int MinRows = 100, MinCols = 26;

        public string PendingOpen;

        Sheet sheet = new Sheet();
        DataGridView grid;
        MenuStrip menu;
        ToolStrip tool;
        StatusStrip status;
        ToolStripStatusLabel lblPos, lblSize, lblSum, lblEnc, lblNl, lblDelim;
        ContextMenuStrip ctx;
        ToolStripMenuItem miUndo, miRedo, miRecent, miFreeze;
        List<ToolStripMenuItem> encItems = new List<ToolStripMenuItem>();
        List<ToolStripMenuItem> delimItems = new List<ToolStripMenuItem>();
        List<ToolStripMenuItem> nlItems = new List<ToolStripMenuItem>();
        List<ToolStripMenuItem> ovfItems = new List<ToolStripMenuItem>();
        ToolStripDropDownButton tbOverflow;
        ToolStripStatusLabel lblZoom;
        ToolStripButton tbUndo, tbRedo, tbSave;

        string filePath;
        EncKind enc = EncKind.Utf8Bom;
        char delim = ',';
        string newline = "\r\n";
        bool headerFrozen;
        bool extraDirty;

        // 表示設定: はみ出し時の扱い
        enum Overflow { Wrap = 0, Spill = 1, Clip = 2 }
        Overflow overflow = Overflow.Spill;
        static readonly string[] OverflowNames = { "折り返す（セル幅で改行）", "はみ出す（右の空セルへ）", "切り詰める" };
        const float DefaultFontSize = 10f;

        List<UndoItem> undoStack = new List<UndoItem>();
        List<UndoItem> redoStack = new List<UndoItem>();
        UndoItem savedTop;
        List<int> pendingWidths;

        Font baseFont;
        Dictionary<FontStyle, Font> fontCache = new Dictionary<FontStyle, Font>();
        float dpi = 96f;
        FindForm findForm;
        Settings settings = new Settings();

        int S(int v) { return (int)Math.Round(v * dpi / 96f); }
        int DefColW { get { return S(90); } }

        public MainForm()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) dpi = g.DpiX;
            Font = SystemFonts.MessageBoxFont;
            Text = AppName;
            AllowDrop = true;
            KeyPreview = false;

            settings.Load();
            string fname = settings.Get("FontName", Font.FontFamily.Name);
            float fsize;
            if (!float.TryParse(settings.Get("FontSize", "10"), NumberStyles.Float, CultureInfo.InvariantCulture, out fsize)) fsize = 10f;
            try { baseFont = new Font(fname, Math.Max(6f, Math.Min(48f, fsize))); } catch { baseFont = new Font(Font.FontFamily, 10f); }

            int ov;
            if (int.TryParse(settings.Get("Overflow", "1"), out ov) && ov >= 0 && ov <= 2) overflow = (Overflow)ov;

            BuildGrid();
            BuildMenu();
            BuildToolbar();
            BuildStatus();
            BuildContextMenu();

            Controls.Add(grid);
            Controls.Add(tool);
            Controls.Add(menu);
            Controls.Add(status);
            MainMenuStrip = menu;

            RestoreWindow();
            ApplyFont(baseFont);
            ApplyOverflow(overflow);
            SyncGrid();

            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (!string.IsNullOrEmpty(PendingOpen)) LoadFile(PendingOpen, null);
            grid.Focus();
        }

        // ---------- UI 構築 ----------
        void BuildGrid()
        {
            grid = new CsvGrid();
            grid.Dock = DockStyle.Fill;
            grid.VirtualMode = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.AllowUserToOrderColumns = false;
            grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = S(24);
            grid.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect;
            grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
            grid.ShowCellToolTips = false;
            grid.BorderStyle = BorderStyle.None;
            grid.BackgroundColor = SystemColors.Window;
            grid.GridColor = Color.FromArgb(218, 220, 224);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 228, 247);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            grid.AllowDrop = true;
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(grid, true, null);

            grid.CellValueNeeded += (s, e) => { e.Value = sheet.Get(e.RowIndex, e.ColumnIndex); };
            grid.CellValuePushed += OnCellValuePushed;
            grid.CellFormatting += OnCellFormatting;
            grid.CellPainting += OnCellPainting;
            grid.RowPostPaint += OnRowPostPaint;
            grid.KeyDown += OnGridKeyDown;
            grid.CurrentCellChanged += (s, e) => UpdateStatus();
            grid.SelectionChanged += (s, e) => UpdateStatus();
            grid.RowHeaderMouseClick += OnRowHeaderClick;
            grid.CellMouseDown += OnCellMouseDown;
            grid.MouseWheel += OnGridWheel;
            grid.DataError += (s, e) => { e.ThrowException = false; };
            grid.DragEnter += OnDragEnter;
            grid.DragDrop += OnDragDrop;
            grid.EditingControlShowing += (s, e) =>
            {
                var tb = e.Control as TextBox;
                if (tb != null) tb.MaxLength = 0;
            };
        }

        ToolStripMenuItem MI(string text, Keys keys, EventHandler h)
        {
            var m = new ToolStripMenuItem(text);
            if (keys != Keys.None) m.ShortcutKeys = keys;
            if (h != null) m.Click += h;
            return m;
        }

        void BuildMenu()
        {
            menu = new MenuStrip();
            menu.Font = Font;

            // ファイル
            var file = new ToolStripMenuItem("ファイル(&F)");
            miRecent = new ToolStripMenuItem("最近使ったファイル(&R)");
            miRecent.DropDownItems.Add("(なし)");
            file.DropDownOpening += (s, e) => RebuildRecent();
            var reload = new ToolStripMenuItem("文字コードを指定して開き直す");
            foreach (EncKind k in new[] { EncKind.Utf8, EncKind.ShiftJis, EncKind.Utf16 })
            {
                EncKind kk = k;
                reload.DropDownItems.Add(MI(Csv.EncName(k), Keys.None, (s, e) => ReloadWith(kk)));
            }
            var encMenu = new ToolStripMenuItem("保存時の文字コード(&E)");
            foreach (EncKind k in new[] { EncKind.Utf8Bom, EncKind.Utf8, EncKind.ShiftJis, EncKind.Utf16 })
            {
                EncKind kk = k;
                var m = MI(Csv.EncName(k), Keys.None, (s, e) => { if (enc != kk) { enc = kk; extraDirty = true; RefreshOptionMenus(); UpdateTitle(); } });
                m.Tag = k; encItems.Add(m); encMenu.DropDownItems.Add(m);
            }
            var delimMenu = new ToolStripMenuItem("区切り文字(&D)");
            foreach (var p in new[] { new KeyValuePair<char, string>(',', "カンマ ( , )"), new KeyValuePair<char, string>('\t', "タブ"), new KeyValuePair<char, string>(';', "セミコロン ( ; )") })
            {
                char cc = p.Key;
                var m = MI(p.Value, Keys.None, (s, e) => { if (delim != cc) { delim = cc; extraDirty = true; RefreshOptionMenus(); UpdateTitle(); } });
                m.Tag = cc; delimItems.Add(m); delimMenu.DropDownItems.Add(m);
            }
            var nlMenu = new ToolStripMenuItem("改行コード(&L)");
            foreach (var p in new[] { new KeyValuePair<string, string>("\r\n", "CRLF (Windows)"), new KeyValuePair<string, string>("\n", "LF (Mac/Linux)") })
            {
                string nn = p.Key;
                var m = MI(p.Value, Keys.None, (s, e) => { if (newline != nn) { newline = nn; extraDirty = true; RefreshOptionMenus(); UpdateTitle(); } });
                m.Tag = nn; nlItems.Add(m); nlMenu.DropDownItems.Add(m);
            }
            file.DropDownItems.AddRange(new ToolStripItem[] {
                MI("新規作成(&N)", Keys.Control | Keys.N, (s, e) => NewDoc()),
                MI("新しいウィンドウ(&W)", Keys.Control | Keys.Shift | Keys.N, (s, e) => { try { Process.Start(Application.ExecutablePath); } catch { } }),
                MI("開く(&O)...", Keys.Control | Keys.O, (s, e) => OpenDialog()),
                miRecent,
                reload,
                new ToolStripSeparator(),
                MI("上書き保存(&S)", Keys.Control | Keys.S, (s, e) => Save()),
                MI("名前を付けて保存(&A)...", Keys.Control | Keys.Shift | Keys.S, (s, e) => SaveAs()),
                new ToolStripSeparator(),
                encMenu, delimMenu, nlMenu,
                new ToolStripSeparator(),
                MI("終了(&X)", Keys.Alt | Keys.F4, (s, e) => Close())
            });

            // 編集
            var edit = new ToolStripMenuItem("編集(&E)");
            miUndo = MI("元に戻す(&U)", Keys.Control | Keys.Z, (s, e) => DoUndo());
            miRedo = MI("やり直し(&R)", Keys.Control | Keys.Y, (s, e) => DoRedo());
            var miDel = MI("削除(&D)", Keys.None, (s, e) => ClearSelected());
            miDel.ShortcutKeyDisplayString = "Del";
            edit.DropDownItems.AddRange(new ToolStripItem[] {
                miUndo, miRedo,
                new ToolStripSeparator(),
                MI("切り取り(&T)", Keys.Control | Keys.X, (s, e) => DoCut()),
                MI("コピー(&C)", Keys.Control | Keys.C, (s, e) => DoCopy()),
                MI("貼り付け(&P)", Keys.Control | Keys.V, (s, e) => DoPaste()),
                miDel,
                MI("すべて選択(&A)", Keys.Control | Keys.A, (s, e) => DoSelectAll()),
                new ToolStripSeparator(),
                MI("下方向へコピー(&W)", Keys.Control | Keys.D, (s, e) => FillDown()),
                MI("右方向へコピー(&I)", Keys.Control | Keys.R, (s, e) => FillRight()),
                new ToolStripSeparator(),
                MI("検索(&F)...", Keys.Control | Keys.F, (s, e) => ShowFind(false)),
                MI("次を検索(&N)", Keys.F3, (s, e) => FindAgain(false)),
                MI("前を検索(&V)", Keys.Shift | Keys.F3, (s, e) => FindAgain(true)),
                MI("置換(&H)...", Keys.Control | Keys.H, (s, e) => ShowFind(true)),
                MI("指定のセルへ移動(&G)...", Keys.Control | Keys.G, (s, e) => GoToCell())
            });

            // 行・列
            var rc = new ToolStripMenuItem("行・列(&R)");
            miFreeze = MI("1行目を見出しとして固定(&H)", Keys.None, (s, e) => ToggleFreeze());
            rc.DropDownItems.AddRange(new ToolStripItem[] {
                MI("上に行を挿入", Keys.Control | Keys.Shift | Keys.Oemplus, (s, e) => InsertRows(false)),
                MI("下に行を挿入", Keys.None, (s, e) => InsertRows(true)),
                MI("選択した行を削除", Keys.Control | Keys.OemMinus, (s, e) => DeleteRows()),
                new ToolStripSeparator(),
                MI("左に列を挿入", Keys.None, (s, e) => InsertCols(false)),
                MI("右に列を挿入", Keys.None, (s, e) => InsertCols(true)),
                MI("選択した列を削除", Keys.None, (s, e) => DeleteCols()),
                new ToolStripSeparator(),
                MI("昇順で並べ替え (A→Z)", Keys.None, (s, e) => SortBy(false)),
                MI("降順で並べ替え (Z→A)", Keys.None, (s, e) => SortBy(true)),
            });
            ((ToolStripMenuItem)rc.DropDownItems[0]).ShortcutKeyDisplayString = "Ctrl++";
            ((ToolStripMenuItem)rc.DropDownItems[2]).ShortcutKeyDisplayString = "Ctrl+-";


            // 表示
            var view = new ToolStripMenuItem("表示(&V)");
            var size = new ToolStripMenuItem("文字サイズ(&S)");
            var zin = MI("拡大", Keys.Control | Keys.Oemplus, (s, e) => Zoom(1));
            zin.ShortcutKeyDisplayString = "Ctrl+; / Ctrl+ホイール";
            var zout = MI("縮小", Keys.Control | Keys.Shift | Keys.OemMinus, (s, e) => Zoom(-1));
            zout.ShortcutKeyDisplayString = "Ctrl+Shift+- / Ctrl+ホイール";
            size.DropDownItems.AddRange(new ToolStripItem[] {
                zin, zout,
                MI("標準のサイズに戻す", Keys.Control | Keys.D0, (s, e) => SetFontSize(DefaultFontSize)),
                new ToolStripSeparator()
            });
            foreach (int pt in new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24 })
            {
                int pp = pt;
                size.DropDownItems.Add(MI(pt + " pt", Keys.None, (s, e) => SetFontSize(pp)));
            }
            size.DropDownOpening += (s, e) =>
            {
                foreach (ToolStripItem it in size.DropDownItems)
                {
                    var m = it as ToolStripMenuItem;
                    if (m != null && m.Text.EndsWith(" pt")) m.Checked = m.Text == (int)Math.Round(baseFont.SizeInPoints) + " pt";
                }
            };
            var ovf = new ToolStripMenuItem("セルからはみ出す文字(&O)");
            for (int i = 0; i < 3; i++)
            {
                var mode = (Overflow)i;
                var m = MI(OverflowNames[i], Keys.None, (s, e) => ApplyOverflow(mode));
                m.Tag = mode; ovfItems.Add(m); ovf.DropDownItems.Add(m);
            }
            view.DropDownItems.AddRange(new ToolStripItem[] {
                size, ovf,
                MI("フォント(&F)...", Keys.None, (s, e) => PickFont()),
                new ToolStripSeparator(),
                miFreeze,
                MI("列幅を自動調整(&A)", Keys.None, (s, e) => AutoFitColumns(true))
            });

            // ヘルプ
            var help = new ToolStripMenuItem("ヘルプ(&H)");
            help.DropDownItems.Add(MI("使い方とショートカット", Keys.F1, (s, e) => ShowHelp()));

            menu.Items.AddRange(new ToolStripItem[] { file, edit, rc, view, help });
            RefreshOptionMenus();
        }

        ToolStripButton TB(string text, string tip, EventHandler h)
        {
            var b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.ToolTipText = tip;
            b.Click += h;
            return b;
        }

        void BuildToolbar()
        {
            tool = new ToolStrip();
            tool.GripStyle = ToolStripGripStyle.Hidden;
            tool.Font = Font;
            tool.Padding = new Padding(S(4), S(2), S(4), S(2));
            tbSave = TB("保存", "上書き保存 (Ctrl+S)", (s, e) => Save());
            tbUndo = TB("↶ 元に戻す", "元に戻す (Ctrl+Z)", (s, e) => DoUndo());
            tbRedo = TB("↷ やり直し", "やり直し (Ctrl+Y)", (s, e) => DoRedo());
            tbOverflow = new ToolStripDropDownButton();
            tbOverflow.DisplayStyle = ToolStripItemDisplayStyle.Text;
            tbOverflow.ToolTipText = "セルからはみ出す文字の表示方法";
            for (int i = 0; i < 3; i++)
            {
                var mode = (Overflow)i;
                var m = new ToolStripMenuItem(OverflowNames[i]);
                m.Tag = mode;
                m.Click += (s, e) => ApplyOverflow(mode);
                tbOverflow.DropDownItems.Add(m);
            }
            tool.Items.AddRange(new ToolStripItem[] {
                TB("新規", "新規作成 (Ctrl+N)", (s, e) => NewDoc()),
                TB("開く", "開く (Ctrl+O)", (s, e) => OpenDialog()),
                tbSave,
                new ToolStripSeparator(),
                tbUndo, tbRedo,
                new ToolStripSeparator(),
                TB("切り取り", "切り取り (Ctrl+X)", (s, e) => DoCut()),
                TB("コピー", "コピー (Ctrl+C)", (s, e) => DoCopy()),
                TB("貼り付け", "貼り付け (Ctrl+V)", (s, e) => DoPaste()),
                new ToolStripSeparator(),
                TB("A+", "文字を大きく (Ctrl+;)", (s, e) => Zoom(1)),
                TB("A-", "文字を小さく (Ctrl+Shift+-)", (s, e) => Zoom(-1)),
                tbOverflow,
                new ToolStripSeparator(),
                TB("昇順", "選択列で昇順に並べ替え", (s, e) => SortBy(false)),
                TB("降順", "選択列で降順に並べ替え", (s, e) => SortBy(true)),
                new ToolStripSeparator(),
                TB("検索", "検索 (Ctrl+F)", (s, e) => ShowFind(false)),
                TB("置換", "置換 (Ctrl+H)", (s, e) => ShowFind(true))
            });
        }

        void BuildStatus()
        {
            status = new StatusStrip();
            status.Font = Font;
            lblPos = new ToolStripStatusLabel();
            lblSize = new ToolStripStatusLabel();
            lblSum = new ToolStripStatusLabel(); lblSum.Spring = true; lblSum.TextAlign = ContentAlignment.MiddleRight;
            lblDelim = new ToolStripStatusLabel();
            lblNl = new ToolStripStatusLabel();
            lblEnc = new ToolStripStatusLabel();
            lblZoom = new ToolStripStatusLabel();
            foreach (var l in new[] { lblPos, lblSize, lblDelim, lblNl, lblEnc, lblZoom })
            {
                l.BorderSides = ToolStripStatusLabelBorderSides.Left;
                l.Padding = new Padding(S(6), 0, S(6), 0);
            }
            lblPos.BorderSides = ToolStripStatusLabelBorderSides.None;
            status.Items.AddRange(new ToolStripItem[] { lblPos, lblSize, lblSum, lblDelim, lblNl, lblEnc, lblZoom });
        }

        void BuildContextMenu()
        {
            ctx = new ContextMenuStrip();
            ctx.Font = Font;
            var del = MI("削除", Keys.None, (s, e) => ClearSelected()); del.ShortcutKeyDisplayString = "Del";
            ctx.Items.AddRange(new ToolStripItem[] {
                MI("切り取り", Keys.None, (s, e) => DoCut()),
                MI("コピー", Keys.None, (s, e) => DoCopy()),
                MI("貼り付け", Keys.None, (s, e) => DoPaste()),
                del,
                new ToolStripSeparator(),
                MI("上に行を挿入", Keys.None, (s, e) => InsertRows(false)),
                MI("下に行を挿入", Keys.None, (s, e) => InsertRows(true)),
                MI("選択した行を削除", Keys.None, (s, e) => DeleteRows()),
                new ToolStripSeparator(),
                MI("左に列を挿入", Keys.None, (s, e) => InsertCols(false)),
                MI("右に列を挿入", Keys.None, (s, e) => InsertCols(true)),
                MI("選択した列を削除", Keys.None, (s, e) => DeleteCols())
            });
            grid.ContextMenuStrip = ctx;
        }

        // ---------- グリッド表示 ----------
        static string ColName(int i)
        {
            string s = "";
            i++;
            while (i > 0) { int m = (i - 1) % 26; s = (char)('A' + m) + s; i = (i - 1) / 26; }
            return s;
        }

        void SyncGrid()
        {
            int wantC = Math.Max(sheet.Cols + 1, MinCols);
            int wantR = Math.Max(sheet.Rows.Count + 1, MinRows);
            if (grid.Columns.Count < wantC)
            {
                var arr = new DataGridViewColumn[wantC - grid.Columns.Count];
                for (int i = 0; i < arr.Length; i++)
                {
                    var col = new DataGridViewTextBoxColumn();
                    col.CellTemplate = new CsvCell();
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                    col.FillWeight = 1f;
                    col.MaxInputLength = 1000000;
                    col.Width = DefColW;
                    col.HeaderText = ColName(grid.Columns.Count + i);
                    arr[i] = col;
                }
                grid.Columns.AddRange(arr);
            }
            else
            {
                while (grid.Columns.Count > wantC) grid.Columns.RemoveAt(grid.Columns.Count - 1);
            }
            if (grid.RowCount != wantR) grid.RowCount = wantR;

            int digits = Math.Max(3, wantR.ToString().Length);
            int w = TextRenderer.MeasureText(new string('8', digits), grid.Font).Width + S(24);
            if (grid.RowHeadersWidth != w) grid.RowHeadersWidth = w;

            if (pendingWidths != null)
            {
                for (int i = 0; i < pendingWidths.Count && i < grid.Columns.Count; i++)
                    grid.Columns[i].Width = Math.Max(S(20), pendingWidths[i]);
                pendingWidths = null;
            }
            ApplyFreeze();
            grid.Invalidate();
            UpdateStatus();
            UpdateTitle();
            UpdateUndoUI();
        }

        void ApplyFreeze()
        {
            if (grid.RowCount > 0) grid.Rows[0].Frozen = headerFrozen;
            miFreeze.Checked = headerFrozen;
        }

        Font GetFont(FontStyle st)
        {
            Font f;
            if (!fontCache.TryGetValue(st, out f)) { f = new Font(baseFont, st); fontCache[st] = f; }
            return f;
        }

        void ApplyFont(Font f)
        {
            baseFont = f;
            foreach (var x in fontCache.Values) x.Dispose();
            fontCache.Clear();
            grid.DefaultCellStyle.Font = f;
            grid.RowTemplate.Height = TextRenderer.MeasureText("Ag漢", f).Height + S(6);
            ResetRows();
            settings.Set("FontName", f.FontFamily.Name);
            settings.Set("FontSize", f.SizeInPoints.ToString(CultureInfo.InvariantCulture));
        }

        void ResetRows()
        {
            if (grid.Columns.Count == 0) return;
            var cur = grid.CurrentCell;
            int cr = cur != null ? cur.RowIndex : 0, cc = cur != null ? cur.ColumnIndex : 0;
            int first = grid.FirstDisplayedScrollingRowIndex;
            grid.RowCount = 0;
            SyncGrid();
            if (cr < grid.RowCount && cc < grid.Columns.Count) grid.CurrentCell = grid[cc, cr];
            if (first >= 0 && first < grid.RowCount) { try { grid.FirstDisplayedScrollingRowIndex = first; } catch { } }
        }

        void OnCellFormatting(object s, DataGridViewCellFormattingEventArgs e)
        {
            if (headerFrozen && e.RowIndex == 0 && e.ColumnIndex >= 0)
            {
                e.CellStyle.Font = GetFont(FontStyle.Bold);
                e.CellStyle.BackColor = Color.FromArgb(243, 244, 246);
            }
        }

        // 「はみ出す」モード: 右隣が空なら文字を右のセルまで描画する (Excel と同じ挙動)
        const TextFormatFlags SpillFlags = TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine |
            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPadding | TextFormatFlags.PreserveGraphicsClipping;

        static string OneLine(string v)
        {
            if (v.IndexOf('\r') < 0 && v.IndexOf('\n') < 0) return v;
            return v.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
        }

        void OnCellPainting(object s, DataGridViewCellPaintingEventArgs e)
        {
            if (overflow != Overflow.Spill || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            int r = e.RowIndex, c = e.ColumnIndex;
            if (grid.IsCurrentCellInEditMode && grid.CurrentCell != null &&
                grid.CurrentCell.RowIndex == r && grid.CurrentCell.ColumnIndex == c) return;

            // 背景・枠線・選択色は標準で描き、文字だけ自前で描く
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);

            int src = c;
            int left = e.CellBounds.Left;
            if (sheet.Get(r, c).Length == 0)
            {
                // 左方向で最初に値のあるセルを探す (30 列まで)
                src = -1;
                int x = e.CellBounds.Left;
                for (int k = c - 1; k >= 0 && k >= c - 30; k--)
                {
                    x -= grid.Columns[k].Width;
                    if (sheet.Get(r, k).Length > 0) { src = k; left = x; break; }
                }
                if (src < 0) { e.Handled = true; return; }
            }
            string text = OneLine(sheet.Get(r, src));
            var font = e.CellStyle.Font ?? baseFont;
            int pad = S(3);
            int textW = TextRenderer.MeasureText(e.Graphics, text, font, new Size(int.MaxValue, int.MaxValue), SpillFlags).Width;

            // 文字が届く右端 (次に値のあるセルの手前で止める)
            int reach = left + pad + textW;
            int spanRight = left + grid.Columns[src].Width;
            for (int k = src + 1; k < grid.Columns.Count && spanRight < reach; k++)
            {
                if (sheet.Get(r, k).Length > 0) break;
                spanRight += grid.Columns[k].Width;
            }
            if (src != c && e.CellBounds.Left >= spanRight) { e.Handled = true; return; }

            bool sel = (e.State & DataGridViewElementStates.Selected) != 0;
            var color = sel ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor;
            var clip = Rectangle.Intersect(e.CellBounds, new Rectangle(left, e.CellBounds.Top, spanRight - left - 1, e.CellBounds.Height - 1));
            if (clip.Width > 0)
            {
                var state = e.Graphics.Save();
                e.Graphics.SetClip(clip);
                var rect = new Rectangle(left + pad, e.CellBounds.Top, textW + pad, e.CellBounds.Height - 1);
                TextRenderer.DrawText(e.Graphics, text, font, rect, color, SpillFlags);
                e.Graphics.Restore(state);
            }
            e.Handled = true;
        }

        static string ShortName(Overflow mode)
        {
            string n = OverflowNames[(int)mode];
            int i = n.IndexOf('（');
            return i > 0 ? n.Substring(0, i) : n;
        }

        void ApplyOverflow(Overflow mode)
        {
            overflow = mode;
            grid.EndEdit();
            grid.DefaultCellStyle.WrapMode = mode == Overflow.Wrap ? DataGridViewTriState.True : DataGridViewTriState.False;
            var newMode = mode == Overflow.Wrap ? DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders : DataGridViewAutoSizeRowsMode.None;
            if (grid.AutoSizeRowsMode != newMode)
            {
                grid.AutoSizeRowsMode = newMode;
                if (newMode == DataGridViewAutoSizeRowsMode.None) ResetRows();   // 行の高さを標準に戻す
            }
            foreach (var m in ovfItems) m.Checked = (Overflow)m.Tag == mode;
            foreach (ToolStripItem it in tbOverflow.DropDownItems)
            {
                var m = it as ToolStripMenuItem;
                if (m != null) m.Checked = (Overflow)m.Tag == mode;
            }
            tbOverflow.Text = "表示: " + ShortName(mode);
            settings.Set("Overflow", ((int)mode).ToString());
            grid.Invalidate();
        }

        void OnRowPostPaint(object s, DataGridViewRowPostPaintEventArgs e)
        {
            var rect = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - S(6), e.RowBounds.Height);
            bool inData = e.RowIndex < sheet.Rows.Count;
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), grid.Font, rect,
                inData ? SystemColors.ControlText : SystemColors.GrayText,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        // ---------- 選択範囲 ----------
        bool AllSelected { get { return grid.Columns.Count > 0 && grid.AreAllCellsSelected(false); } }

        Rectangle SelBounds()
        {
            if (AllSelected) return new Rectangle(0, 0, Math.Max(sheet.Cols, 1), Math.Max(sheet.Rows.Count, 1));
            int minR = int.MaxValue, minC = int.MaxValue, maxR = -1, maxC = -1;
            foreach (DataGridViewCell c in grid.SelectedCells)
            {
                if (c.RowIndex < minR) minR = c.RowIndex;
                if (c.RowIndex > maxR) maxR = c.RowIndex;
                if (c.ColumnIndex < minC) minC = c.ColumnIndex;
                if (c.ColumnIndex > maxC) maxC = c.ColumnIndex;
            }
            if (maxR < 0)
            {
                var cur = grid.CurrentCell;
                if (cur == null) return Rectangle.Empty;
                return new Rectangle(cur.ColumnIndex, cur.RowIndex, 1, 1);
            }
            return Rectangle.FromLTRB(minC, minR, maxC + 1, maxR + 1);
        }

        // 選択セルを (X=列, Y=行) で列挙
        IEnumerable<Point> SelPoints()
        {
            if (AllSelected)
            {
                for (int r = 0; r < sheet.Rows.Count; r++)
                    for (int c = 0; c < sheet.Cols; c++) yield return new Point(c, r);
                yield break;
            }
            if (grid.SelectedCells.Count == 0)
            {
                var cur = grid.CurrentCell;
                if (cur != null) yield return new Point(cur.ColumnIndex, cur.RowIndex);
                yield break;
            }
            foreach (DataGridViewCell c in grid.SelectedCells) yield return new Point(c.ColumnIndex, c.RowIndex);
        }

        void SelectRange(int c1, int r1, int c2, int r2)
        {
            c2 = Math.Min(c2, grid.Columns.Count - 1);
            r2 = Math.Min(r2, grid.RowCount - 1);
            if (c1 > c2 || r1 > r2) return;
            grid.CurrentCell = grid[c1, r1];
            if (c1 == c2 && r1 == r2) return;
            long cells = (long)(c2 - c1 + 1) * (r2 - r1 + 1);
            if (cells > 200000) return;
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++) grid[c, r].Selected = true;
        }

        int rowAnchor = -1;
        void OnRowHeaderClick(object s, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            grid.EndEdit();
            int lastC = Math.Max(sheet.Cols, 1) - 1;
            if ((ModifierKeys & Keys.Shift) != 0 && rowAnchor >= 0)
            {
                int a = Math.Min(rowAnchor, e.RowIndex), b = Math.Max(rowAnchor, e.RowIndex);
                SelectRange(0, a, lastC, b);
            }
            else
            {
                rowAnchor = e.RowIndex;
                grid.ClearSelection();
                SelectRange(0, e.RowIndex, lastC, e.RowIndex);
            }
        }

        void OnCellMouseDown(object s, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var cell = grid[e.ColumnIndex, e.RowIndex];
                if (!cell.Selected) { grid.EndEdit(); grid.CurrentCell = cell; }
            }
        }

        void OnGridWheel(object s, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) != 0)
            {
                Zoom(e.Delta > 0 ? 1 : -1);
                var he = e as HandledMouseEventArgs;
                if (he != null) he.Handled = true;
            }
        }

        void OnGridKeyDown(object s, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.None)
            {
                ClearSelected(); e.Handled = true;
            }
            else if (e.KeyCode == Keys.Back && e.Modifiers == Keys.None && grid.CurrentCell != null)
            {
                grid.BeginEdit(true);
                var tb = EditBox;
                if (tb != null) tb.Text = "";
                e.Handled = true;
            }
        }

        TextBox EditBox
        {
            get { return grid.IsCurrentCellInEditMode ? grid.EditingControl as TextBox : null; }
        }

        bool GridHasFocus { get { return grid.ContainsFocus; } }

        // ---------- 元に戻す / やり直し ----------
        void Push(UndoItem u)
        {
            undoStack.Add(u);
            if (undoStack.Count > 1000) undoStack.RemoveAt(0);
            redoStack.Clear();
        }

        void DoUndo()
        {
            var tb = EditBox;
            if (tb != null) { if (tb.CanUndo) tb.Undo(); else grid.CancelEdit(); return; }
            if (undoStack.Count == 0) return;
            var u = undoStack[undoStack.Count - 1];
            undoStack.RemoveAt(undoStack.Count - 1);
            u.Undo();
            redoStack.Add(u);
            AfterChange();
        }

        void DoRedo()
        {
            if (EditBox != null) return;
            if (redoStack.Count == 0) return;
            var u = redoStack[redoStack.Count - 1];
            redoStack.RemoveAt(redoStack.Count - 1);
            u.Redo();
            undoStack.Add(u);
            AfterChange();
        }

        void AfterChange() { SyncGrid(); }

        bool IsDirty
        {
            get
            {
                var top = undoStack.Count > 0 ? undoStack[undoStack.Count - 1] : null;
                return extraDirty || top != savedTop;
            }
        }

        void MarkSaved()
        {
            savedTop = undoStack.Count > 0 ? undoStack[undoStack.Count - 1] : null;
            extraDirty = false;
            UpdateTitle();
        }

        void UpdateUndoUI()
        {
            miUndo.Enabled = tbUndo.Enabled = undoStack.Count > 0;
            miRedo.Enabled = tbRedo.Enabled = redoStack.Count > 0;
            miUndo.Text = undoStack.Count > 0 ? "元に戻す: " + undoStack[undoStack.Count - 1].Name + "(&U)" : "元に戻す(&U)";
            miRedo.Text = redoStack.Count > 0 ? "やり直し: " + redoStack[redoStack.Count - 1].Name + "(&R)" : "やり直し(&R)";
        }

        void ApplyCellChanges(List<CellChange> ch, string name, bool defer)
        {
            if (ch.Count == 0) return;
            int oR = sheet.Rows.Count, oC = sheet.Cols;
            Action redo = () => { foreach (var x in ch) sheet.Set(x.R, x.C, x.New); };
            Action undo = () =>
            {
                foreach (var x in ch)
                    if (x.R < sheet.Rows.Count && x.C < sheet.Cols) sheet.Rows[x.R].V[x.C] = x.Old;
                sheet.Truncate(oR, oC);
            };
            redo();
            Push(new UndoItem { Name = name, Undo = undo, Redo = redo });
            if (defer) BeginInvoke((Action)AfterChange); else AfterChange();
        }

        void OnCellValuePushed(object s, DataGridViewCellValueEventArgs e)
        {
            string nv = e.Value == null ? "" : e.Value.ToString();
            string old = sheet.Get(e.RowIndex, e.ColumnIndex);
            if (nv == old) return;
            ApplyCellChanges(new List<CellChange> { new CellChange { R = e.RowIndex, C = e.ColumnIndex, Old = old, New = nv } }, "入力", true);
        }

        // ---------- クリップボード ----------
        static string EscapeTsv(string s)
        {
            if (s.IndexOf('\t') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0 || s.IndexOf('"') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        bool CopySelection()
        {
            var b = SelBounds();
            if (b.Width <= 0 || b.Height <= 0) return false;
            if ((long)b.Width * b.Height > 5000000)
            {
                MessageBox.Show(this, "選択範囲が大きすぎるためコピーできません。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            var sb = new StringBuilder();
            for (int r = b.Top; r < b.Bottom; r++)
            {
                for (int c = b.Left; c < b.Right; c++)
                {
                    if (c > b.Left) sb.Append('\t');
                    sb.Append(EscapeTsv(sheet.Get(r, c)));
                }
                if (r < b.Bottom - 1 || b.Height > 1) sb.Append("\r\n");
            }
            try
            {
                var d = new DataObject();
                d.SetData(DataFormats.UnicodeText, sb.ToString());
                d.SetData(DataFormats.Text, sb.ToString());
                Clipboard.SetDataObject(d, true, 10, 100);
                return true;
            }
            catch (Exception)
            {
                // クリップボードが他のアプリに使われている等で失敗しても通知はしない。
                // 実際には同じ内容が入っていることが多いので、それなら成功扱いにする (切り取りのため)
                try { return Clipboard.GetText(TextDataFormat.UnicodeText) == sb.ToString(); }
                catch { return false; }
            }
        }

        void DoCopy()
        {
            if (!GridHasFocus) { var tb0 = ActiveControl as TextBoxBase; if (tb0 != null) tb0.Copy(); return; }
            var tb = EditBox;
            if (tb != null) { tb.Copy(); return; }
            CopySelection();
        }

        void DoCut()
        {
            if (!GridHasFocus) { var tb0 = ActiveControl as TextBoxBase; if (tb0 != null) tb0.Cut(); return; }
            var tb = EditBox;
            if (tb != null) { tb.Cut(); return; }
            if (CopySelection()) ClearSelected("切り取り");
        }

        void DoPaste()
        {
            if (!GridHasFocus) { var tb0 = ActiveControl as TextBoxBase; if (tb0 != null) tb0.Paste(); return; }
            var tb = EditBox;
            if (tb != null) { tb.Paste(); return; }
            string text;
            try { text = Clipboard.GetText(TextDataFormat.UnicodeText); }
            catch { return; }
            if (string.IsNullOrEmpty(text)) return;
            var data = Csv.Parse(text, '\t');
            if (data.Count == 0) return;
            var b = SelBounds();
            if (b.Width <= 0) return;
            var ch = new List<CellChange>();
            int selCount = AllSelected ? 2 : grid.GetCellCount(DataGridViewElementStates.Selected);
            if (data.Count == 1 && data[0].Count == 1 && selCount > 1)
            {
                // 1つの値を選択範囲すべてに貼り付け
                string v = data[0][0];
                foreach (var p in SelPoints())
                {
                    string old = sheet.Get(p.Y, p.X);
                    if (old != v) ch.Add(new CellChange { R = p.Y, C = p.X, Old = old, New = v });
                }
                ApplyCellChanges(ch, "貼り付け", false);
                return;
            }
            int maxW = 0;
            for (int i = 0; i < data.Count; i++)
            {
                maxW = Math.Max(maxW, data[i].Count);
                for (int j = 0; j < data[i].Count; j++)
                {
                    int r = b.Top + i, c = b.Left + j;
                    string old = sheet.Get(r, c);
                    if (old != data[i][j]) ch.Add(new CellChange { R = r, C = c, Old = old, New = data[i][j] });
                }
            }
            ApplyCellChanges(ch, "貼り付け", false);
            grid.ClearSelection();
            SelectRange(b.Left, b.Top, b.Left + maxW - 1, b.Top + data.Count - 1);
        }

        void DoSelectAll()
        {
            if (!GridHasFocus) { var tb0 = ActiveControl as TextBoxBase; if (tb0 != null) tb0.SelectAll(); return; }
            var tb = EditBox;
            if (tb != null) { tb.SelectAll(); return; }
            grid.SelectAll();
        }

        // Ctrl+D: 下方向へコピー (1行だけ選択中なら上のセルをコピー)
        void FillDown()
        {
            grid.EndEdit();
            var b = SelBounds();
            if (b.Width <= 0 || b.Height <= 0) return;
            int src = b.Height == 1 ? b.Top - 1 : b.Top;
            if (src < 0) return;
            var ch = new List<CellChange>();
            for (int c = b.Left; c < b.Right; c++)
            {
                string v = sheet.Get(src, c);
                for (int r = Math.Max(src + 1, b.Top); r < b.Bottom; r++)
                {
                    string old = sheet.Get(r, c);
                    if (old != v) ch.Add(new CellChange { R = r, C = c, Old = old, New = v });
                }
            }
            ApplyCellChanges(ch, "下方向へコピー", false);
        }

        // Ctrl+R: 右方向へコピー (1列だけ選択中なら左のセルをコピー)
        void FillRight()
        {
            grid.EndEdit();
            var b = SelBounds();
            if (b.Width <= 0 || b.Height <= 0) return;
            int src = b.Width == 1 ? b.Left - 1 : b.Left;
            if (src < 0) return;
            var ch = new List<CellChange>();
            for (int r = b.Top; r < b.Bottom; r++)
            {
                string v = sheet.Get(r, src);
                for (int c = Math.Max(src + 1, b.Left); c < b.Right; c++)
                {
                    string old = sheet.Get(r, c);
                    if (old != v) ch.Add(new CellChange { R = r, C = c, Old = old, New = v });
                }
            }
            ApplyCellChanges(ch, "右方向へコピー", false);
        }

        void ClearSelected() { ClearSelected("削除"); }
        void ClearSelected(string name)
        {
            if (EditBox != null) return;
            var ch = new List<CellChange>();
            foreach (var p in SelPoints())
            {
                string old = sheet.Get(p.Y, p.X);
                if (old.Length > 0) ch.Add(new CellChange { R = p.Y, C = p.X, Old = old, New = "" });
            }
            ApplyCellChanges(ch, name, false);
        }

        // ---------- 行・列操作 ----------
        List<int> GetWidths()
        {
            var w = new List<int>(grid.Columns.Count);
            foreach (DataGridViewColumn c in grid.Columns) w.Add(c.Width);
            return w;
        }

        void InsertRows(bool below)
        {
            grid.EndEdit();
            var b = SelBounds();
            if (b.Height <= 0) return;
            int n = Math.Min(b.Height, Math.Max(1, sheet.Rows.Count));
            int at = Math.Min(below ? b.Bottom : b.Top, sheet.Rows.Count);
            if (sheet.Cols == 0) sheet.EnsureSize(0, 1);
            Action redo = () =>
            {
                var list = new List<Row>();
                for (int i = 0; i < n; i++) list.Add(new Row(sheet.Cols));
                sheet.Rows.InsertRange(at, list);
            };
            Action undo = () => sheet.Rows.RemoveRange(at, n);
            redo();
            Push(new UndoItem { Name = "行の挿入", Undo = undo, Redo = redo });
            AfterChange();
            SelectRange(0, at, Math.Max(sheet.Cols, 1) - 1, at + n - 1);
        }

        void DeleteRows()
        {
            grid.EndEdit();
            var set = new HashSet<int>();
            foreach (var p in SelPoints()) if (p.Y < sheet.Rows.Count) set.Add(p.Y);
            if (set.Count == 0) return;
            var idx = set.OrderBy(x => x).ToList();
            var removed = idx.Select(i => sheet.Rows[i]).ToList();
            Action redo = () => { for (int k = idx.Count - 1; k >= 0; k--) sheet.Rows.RemoveAt(idx[k]); };
            Action undo = () => { for (int k = 0; k < idx.Count; k++) sheet.Rows.Insert(idx[k], removed[k]); };
            redo();
            Push(new UndoItem { Name = "行の削除", Undo = undo, Redo = redo });
            AfterChange();
        }

        void InsertCols(bool right)
        {
            grid.EndEdit();
            var b = SelBounds();
            if (b.Width <= 0) return;
            int n = Math.Min(b.Width, Math.Max(1, sheet.Cols));
            int at = Math.Min(right ? b.Right : b.Left, sheet.Cols);
            Action redo = () =>
            {
                var w = GetWidths();
                foreach (var row in sheet.Rows)
                    for (int i = 0; i < n; i++) row.V.Insert(at, "");
                sheet.Cols += n;
                for (int i = 0; i < n; i++) w.Insert(Math.Min(at, w.Count), DefColW);
                pendingWidths = w;
            };
            Action undo = () =>
            {
                var w = GetWidths();
                foreach (var row in sheet.Rows) row.V.RemoveRange(at, n);
                sheet.Cols -= n;
                if (at + n <= w.Count) w.RemoveRange(at, n);
                pendingWidths = w;
            };
            redo();
            Push(new UndoItem { Name = "列の挿入", Undo = undo, Redo = redo });
            AfterChange();
            SelectRange(at, 0, at + n - 1, Math.Max(sheet.Rows.Count, 1) - 1);
        }

        void DeleteCols()
        {
            grid.EndEdit();
            var set = new HashSet<int>();
            foreach (var p in SelPoints()) if (p.X < sheet.Cols) set.Add(p.X);
            if (set.Count == 0) return;
            var idx = set.OrderBy(x => x).ToList();
            var savedV = new List<string[]>();
            List<int> savedW = null;
            Action redo = () =>
            {
                var w = GetWidths();
                savedW = idx.Select(i => i < w.Count ? w[i] : DefColW).ToList();
                savedV.Clear();
                foreach (var row in sheet.Rows)
                {
                    var v = new string[idx.Count];
                    for (int k = idx.Count - 1; k >= 0; k--) { v[k] = row.V[idx[k]]; row.V.RemoveAt(idx[k]); }
                    savedV.Add(v);
                }
                for (int k = idx.Count - 1; k >= 0; k--) if (idx[k] < w.Count) w.RemoveAt(idx[k]);
                sheet.Cols -= idx.Count;
                pendingWidths = w;
            };
            Action undo = () =>
            {
                var w = GetWidths();
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    var row = sheet.Rows[r];
                    for (int k = 0; k < idx.Count; k++) row.V.Insert(idx[k], savedV[r][k]);
                }
                sheet.Cols += idx.Count;
                for (int k = 0; k < idx.Count; k++) w.Insert(Math.Min(idx[k], w.Count), savedW[k]);
                pendingWidths = w;
            };
            redo();
            Push(new UndoItem { Name = "列の削除", Undo = undo, Redo = redo });
            AfterChange();
        }

        void SortBy(bool desc)
        {
            grid.EndEdit();
            var cur = grid.CurrentCell;
            if (cur == null) return;
            int col = cur.ColumnIndex;
            if (col >= sheet.Cols) return;
            int start = headerFrozen ? 1 : 0;
            if (sheet.Rows.Count - start < 2) return;
            var before = new List<Row>(sheet.Rows);
            var cmp = new SmartComparer();
            var body = sheet.Rows.Skip(start).OrderBy(r => r.V[col].Length == 0 ? 1 : 0);
            var sorted = desc ? body.ThenByDescending(r => r.V[col], cmp) : body.ThenBy(r => r.V[col], cmp);
            var after = sheet.Rows.Take(start).Concat(sorted).ToList();
            Action redo = () => { sheet.Rows.Clear(); sheet.Rows.AddRange(after); };
            Action undo = () => { sheet.Rows.Clear(); sheet.Rows.AddRange(before); };
            redo();
            Push(new UndoItem { Name = desc ? "降順で並べ替え" : "昇順で並べ替え", Undo = undo, Redo = redo });
            AfterChange();
        }

        void ToggleFreeze()
        {
            headerFrozen = !headerFrozen;
            extraDirty = true;
            ApplyFreeze();
            grid.Invalidate();
            UpdateTitle();
        }

        void AutoFitColumns(bool all)
        {
            int n = Math.Min(sheet.Cols, grid.Columns.Count);
            int rows = Math.Min(sheet.Rows.Count, 1000);
            var bold = GetFont(FontStyle.Bold);
            for (int c = 0; c < n; c++)
            {
                int w = S(64);
                for (int r = 0; r < rows; r++)
                {
                    string s = sheet.Rows[r].V[c];
                    if (s.Length == 0) continue;
                    int nl = s.IndexOfAny(new[] { '\r', '\n' });
                    if (nl >= 0) s = s.Substring(0, nl);
                    if (s.Length > 200) s = s.Substring(0, 200);
                    var font = headerFrozen && r == 0 ? bold : baseFont;
                    int tw = TextRenderer.MeasureText(s, font).Width + S(12);
                    if (tw > w) w = tw;
                }
                grid.Columns[c].Width = Math.Min(w, S(400));
            }
        }

        // ---------- 表示設定 ----------
        void PickFont()
        {
            using (var d = new FontDialog())
            {
                d.Font = baseFont;
                d.ShowEffects = false;
                d.AllowVerticalFonts = false;
                if (d.ShowDialog(this) == DialogResult.OK)
                    ApplyFont(new Font(d.Font.FontFamily, d.Font.SizeInPoints, FontStyle.Regular));
            }
        }

        void Zoom(int step)
        {
            SetFontSize((float)Math.Round(baseFont.SizeInPoints) + step);
        }

        void SetFontSize(float size)
        {
            if (size < 6 || size > 48) return;
            ApplyFont(new Font(baseFont.FontFamily, size, FontStyle.Regular));
            UpdateStatus();
        }

        // ---------- ファイル ----------
        bool ConfirmDiscard()
        {
            grid.EndEdit();
            if (!IsDirty) return true;
            var r = MessageBox.Show(this, "「" + DocName + "」への変更を保存しますか?", AppName,
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.Yes) return Save();
            return true;
        }

        string DocName { get { return filePath == null ? "無題" : Path.GetFileName(filePath); } }

        void ResetDoc()
        {
            undoStack.Clear(); redoStack.Clear(); savedTop = null; extraDirty = false;
            grid.RowCount = 0;
            while (grid.Columns.Count > 0) grid.Columns.RemoveAt(grid.Columns.Count - 1);
        }

        void NewDoc()
        {
            if (!ConfirmDiscard()) return;
            sheet = new Sheet();
            filePath = null; enc = EncKind.Utf8Bom; delim = ','; newline = "\r\n"; headerFrozen = false;
            ResetDoc();
            SyncGrid();
            RefreshOptionMenus();
        }

        void OpenDialog()
        {
            if (!ConfirmDiscard()) return;
            using (var d = new OpenFileDialog())
            {
                d.Filter = "CSV / TSV ファイル (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt|すべてのファイル (*.*)|*.*";
                if (filePath != null) d.InitialDirectory = Path.GetDirectoryName(filePath);
                if (d.ShowDialog(this) == DialogResult.OK) LoadFile(d.FileName, null);
            }
        }

        void ReloadWith(EncKind k)
        {
            if (filePath == null || !File.Exists(filePath))
            {
                MessageBox.Show(this, "開き直すファイルがありません。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!ConfirmDiscard()) return;
            LoadFile(filePath, k);
        }

        // ショートカット (.lnk) ならリンク先のパスを返す
        static string ResolveShortcut(string path)
        {
            if (path == null || !path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)) return path;
            object sh = null, lnk = null;
            try
            {
                var t = Type.GetTypeFromProgID("WScript.Shell");
                if (t == null) return path;
                sh = Activator.CreateInstance(t);
                lnk = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, sh, new object[] { path });
                var target = lnk.GetType().InvokeMember("TargetPath", BindingFlags.GetProperty, null, lnk, null) as string;
                if (!string.IsNullOrEmpty(target)) return target;
            }
            catch { }
            finally
            {
                try { if (lnk != null) Marshal.FinalReleaseComObject(lnk); } catch { }
                try { if (sh != null) Marshal.FinalReleaseComObject(sh); } catch { }
            }
            return path;
        }

        // テキストではない (バイナリ) ファイルらしいか
        static bool LooksBinary(byte[] b, EncKind k)
        {
            if (k == EncKind.Utf16) return false;
            int n = Math.Min(b.Length, 8192);
            for (int i = 0; i < n; i++) if (b[i] == 0) return true;
            return false;
        }

        public void LoadFile(string path, EncKind? forced)
        {
            path = ResolveShortcut(path);
            if (Directory.Exists(path))
            {
                using (var dlg = new OpenFileDialog())
                {
                    dlg.Filter = "CSV / TSV ファイル (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt|すべてのファイル (*.*)|*.*";
                    dlg.InitialDirectory = path;
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    path = dlg.FileName;
                }
            }
            Sheet ns;
            EncKind k; char d; string nl;
            try
            {
                Cursor = Cursors.WaitCursor;
                byte[] bytes = File.ReadAllBytes(path);
                if (forced == null)
                {
                    EncKind tmpK;
                    Csv.Decode(bytes.Length > 8192 ? bytes.Take(8192).ToArray() : bytes, null, out tmpK);
                    if (LooksBinary(bytes, tmpK))
                    {
                        Cursor = Cursors.Default;
                        var ans = MessageBox.Show(this,
                            "「" + Path.GetFileName(path) + "」は CSV などのテキストファイルではないようです。\n" +
                            "開くと文字化けして表示されます。それでも開きますか？",
                            AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                        if (ans != DialogResult.Yes) return;
                        Cursor = Cursors.WaitCursor;
                    }
                }
                string text = Csv.Decode(bytes, forced, out k);
                d = Csv.DetectDelimiter(text, path);
                nl = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0 ? "\r\n" : (text.IndexOf('\n') >= 0 ? "\n" : "\r\n");
                var data = Csv.Parse(text, d);
                ns = new Sheet();
                ns.Rows.Capacity = data.Count;
                foreach (var r in data) ns.Rows.Add(new Row(r));
                ns.Normalize();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show(this, "ファイルを開けませんでした。\n" + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            sheet = ns; filePath = path; enc = k; delim = d; newline = nl; headerFrozen = false;
            ResetDoc();
            SyncGrid();
            AutoFitColumns(true);
            ApplyFreeze();
            RefreshOptionMenus();
            AddRecent(path);
            MarkSaved();
            Cursor = Cursors.Default;
            if (grid.RowCount > 0 && grid.Columns.Count > 0) grid.CurrentCell = grid[0, 0];
        }

        bool Save()
        {
            grid.EndEdit();
            if (filePath == null) return SaveAs();
            return WriteFile(filePath);
        }

        bool SaveAs()
        {
            grid.EndEdit();
            using (var d = new SaveFileDialog())
            {
                d.Filter = "CSV ファイル (*.csv)|*.csv|TSV ファイル (*.tsv)|*.tsv|テキスト ファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*";
                d.FilterIndex = delim == '\t' ? 2 : 1;
                d.FileName = filePath == null ? "無題.csv" : Path.GetFileName(filePath);
                if (filePath != null) d.InitialDirectory = Path.GetDirectoryName(filePath);
                d.OverwritePrompt = true;
                if (d.ShowDialog(this) != DialogResult.OK) return false;
                string ext = Path.GetExtension(d.FileName).ToLowerInvariant();
                if (ext == ".tsv") delim = '\t';
                else if (ext == ".csv" && delim == '\t') delim = ',';
                if (!WriteFile(d.FileName)) return false;
                filePath = d.FileName;
                AddRecent(filePath);
                RefreshOptionMenus();
                UpdateTitle();
                return true;
            }
        }

        bool WriteFile(string path)
        {
            var sb = new StringBuilder();
            foreach (var row in sheet.Rows)
            {
                for (int c = 0; c < sheet.Cols; c++)
                {
                    if (c > 0) sb.Append(delim);
                    sb.Append(Csv.Escape(row.V[c], delim));
                }
                sb.Append(newline);
            }
            string text = sb.ToString();
            byte[] bytes;
            try { bytes = Csv.Encode(text, enc, true); }
            catch (EncoderFallbackException)
            {
                var r = MessageBox.Show(this,
                    "Shift_JIS で表せない文字（絵文字・一部の漢字など）が含まれています。\n\n" +
                    "[はい] UTF-8 (BOM付き) に変更して保存\n[いいえ] そのまま保存（該当文字は ? になります）",
                    AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (r == DialogResult.Cancel) return false;
                if (r == DialogResult.Yes) enc = EncKind.Utf8Bom;
                bytes = Csv.Encode(text, enc, false);
            }
            try
            {
                string tmp = path + ".saving~";
                File.WriteAllBytes(tmp, bytes);
                if (File.Exists(path))
                {
                    try { File.Replace(tmp, path, null); }
                    catch { File.Copy(tmp, path, true); File.Delete(tmp); }
                }
                else File.Move(tmp, path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存できませんでした。\n" + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            MarkSaved();
            RefreshOptionMenus();
            return true;
        }

        void RefreshOptionMenus()
        {
            foreach (var m in encItems) m.Checked = (EncKind)m.Tag == enc;
            foreach (var m in delimItems) m.Checked = (char)m.Tag == delim;
            foreach (var m in nlItems) m.Checked = (string)m.Tag == newline;
            UpdateStatus();
        }

        void UpdateTitle()
        {
            string t = (IsDirty ? "*" : "") + DocName + " - " + AppName;
            if (Text != t) Text = t;
        }

        void UpdateStatus()
        {
            if (lblPos == null) return;
            var cur = grid.CurrentCell;
            lblPos.Text = cur != null ? ColName(cur.ColumnIndex) + (cur.RowIndex + 1) : "";
            lblSize.Text = sheet.Rows.Count + " 行 × " + sheet.Cols + " 列";
            lblEnc.Text = Csv.EncName(enc);
            lblZoom.Text = "文字 " + baseFont.SizeInPoints.ToString("0.#") + "pt";
            lblNl.Text = newline == "\r\n" ? "CRLF" : "LF";
            lblDelim.Text = delim == ',' ? "カンマ区切り" : delim == '\t' ? "タブ区切り" : "セミコロン区切り";

            // 選択範囲の個数・合計（Excel 風）
            string sum = "";
            int cnt = AllSelected ? int.MaxValue : grid.GetCellCount(DataGridViewElementStates.Selected);
            if (cnt > 1 && cnt <= 50000)
            {
                int filled = 0, nums = 0; double total = 0;
                foreach (var p in SelPoints())
                {
                    string v = sheet.Get(p.Y, p.X);
                    if (v.Length == 0) continue;
                    filled++;
                    double d;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) { nums++; total += d; }
                }
                sum = "データの個数: " + filled;
                if (nums > 0) sum += "   合計: " + total.ToString("#,0.##########") + "   平均: " + (total / nums).ToString("#,0.####");
            }
            lblSum.Text = sum;
        }

        // ---------- 検索・置換 ----------
        public string LastFind = "";
        public bool LastCase, LastWhole;

        void ShowFind(bool replace)
        {
            grid.EndEdit();
            if (findForm == null || findForm.IsDisposed) findForm = new FindForm(this);
            var cur = grid.CurrentCell;
            if (cur != null && findForm.FindBox.Text.Length == 0)
            {
                string v = sheet.Get(cur.RowIndex, cur.ColumnIndex);
                if (v.Length > 0 && v.Length < 100 && v.IndexOf('\n') < 0) findForm.FindBox.Text = v;
            }
            findForm.SetReplaceMode(replace);
            if (!findForm.Visible)
            {
                var p = PointToScreen(new Point(ClientSize.Width - findForm.Width - S(30), S(80)));
                findForm.Location = p;
                findForm.Show(this);
            }
            findForm.Activate();
            findForm.FindBox.Focus();
            findForm.FindBox.SelectAll();
        }

        void FindAgain(bool back)
        {
            if (string.IsNullOrEmpty(LastFind)) { ShowFind(false); return; }
            if (!FindNext(LastFind, back, LastCase, LastWhole))
                MessageBox.Show(this, "「" + LastFind + "」は見つかりませんでした。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static bool Match(string v, string q, StringComparison cmp, bool whole)
        {
            return whole ? string.Equals(v, q, cmp) : v.IndexOf(q, cmp) >= 0;
        }

        public bool FindNext(string q, bool back, bool matchCase, bool whole)
        {
            grid.EndEdit();
            LastFind = q; LastCase = matchCase; LastWhole = whole;
            if (string.IsNullOrEmpty(q) || sheet.Rows.Count == 0 || sheet.Cols == 0) return false;
            var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int R = sheet.Rows.Count, C = sheet.Cols;
            long total = (long)R * C;
            var cur = grid.CurrentCell;
            long start;
            if (cur != null && cur.RowIndex < R && cur.ColumnIndex < C) start = (long)cur.RowIndex * C + cur.ColumnIndex;
            else start = back ? total : -1;
            for (long k = 1; k <= total; k++)
            {
                long idx = back ? start - k : start + k;
                idx = ((idx % total) + total) % total;
                int r = (int)(idx / C), c = (int)(idx % C);
                if (Match(sheet.Rows[r].V[c], q, cmp, whole))
                {
                    grid.CurrentCell = grid[c, r];
                    return true;
                }
            }
            return false;
        }

        static string ReplaceIn(string v, string q, string rep, StringComparison cmp, bool whole)
        {
            if (whole) return rep;
            var sb = new StringBuilder();
            int pos = 0;
            while (true)
            {
                int i = v.IndexOf(q, pos, cmp);
                if (i < 0) break;
                sb.Append(v, pos, i - pos).Append(rep);
                pos = i + q.Length;
            }
            sb.Append(v, pos, v.Length - pos);
            return sb.ToString();
        }

        public void ReplaceOne(string q, string rep, bool matchCase, bool whole)
        {
            grid.EndEdit();
            if (string.IsNullOrEmpty(q)) return;
            var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var cur = grid.CurrentCell;
            if (cur != null && cur.RowIndex < sheet.Rows.Count && cur.ColumnIndex < sheet.Cols)
            {
                string v = sheet.Get(cur.RowIndex, cur.ColumnIndex);
                if (Match(v, q, cmp, whole))
                {
                    string nv = ReplaceIn(v, q, rep, cmp, whole);
                    if (nv != v)
                        ApplyCellChanges(new List<CellChange> { new CellChange { R = cur.RowIndex, C = cur.ColumnIndex, Old = v, New = nv } }, "置換", false);
                }
            }
            FindNext(q, false, matchCase, whole);
        }

        public int ReplaceAll(string q, string rep, bool matchCase, bool whole, bool selectionOnly)
        {
            grid.EndEdit();
            if (string.IsNullOrEmpty(q)) return 0;
            var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var ch = new List<CellChange>();
            IEnumerable<Point> pts;
            if (selectionOnly) pts = SelPoints();
            else pts = AllCells();
            foreach (var p in pts)
            {
                if (p.Y >= sheet.Rows.Count || p.X >= sheet.Cols) continue;
                string v = sheet.Rows[p.Y].V[p.X];
                if (!Match(v, q, cmp, whole)) continue;
                string nv = ReplaceIn(v, q, rep, cmp, whole);
                if (nv != v) ch.Add(new CellChange { R = p.Y, C = p.X, Old = v, New = nv });
            }
            ApplyCellChanges(ch, "すべて置換", false);
            return ch.Count;
        }

        IEnumerable<Point> AllCells()
        {
            for (int r = 0; r < sheet.Rows.Count; r++)
                for (int c = 0; c < sheet.Cols; c++) yield return new Point(c, r);
        }

        void GoToCell()
        {
            grid.EndEdit();
            string s = InputBox.Show(this, "移動先のセル（例: B12）または行番号を入力してください。", "指定のセルへ移動", "");
            if (string.IsNullOrEmpty(s)) return;
            s = s.Trim().ToUpperInvariant();
            int i = 0, col = 0;
            while (i < s.Length && s[i] >= 'A' && s[i] <= 'Z') { col = col * 26 + (s[i] - 'A' + 1); i++; }
            int row;
            if (!int.TryParse(s.Substring(i), out row) || row < 1) return;
            int c = col == 0 ? (grid.CurrentCell != null ? grid.CurrentCell.ColumnIndex : 0) : col - 1;
            int r = row - 1;
            r = Math.Min(r, grid.RowCount - 1); c = Math.Min(c, grid.Columns.Count - 1);
            grid.CurrentCell = grid[c, r];
        }

        // ---------- その他 ----------
        void RebuildRecent()
        {
            miRecent.DropDownItems.Clear();
            var list = settings.Recent;
            if (list.Count == 0) { var m = new ToolStripMenuItem("(なし)"); m.Enabled = false; miRecent.DropDownItems.Add(m); return; }
            for (int i = 0; i < list.Count; i++)
            {
                string p = list[i];
                miRecent.DropDownItems.Add(MI("&" + ((i + 1) % 10) + "  " + p, Keys.None, (s, e) =>
                {
                    if (!File.Exists(p)) { MessageBox.Show(this, "ファイルが見つかりません。\n" + p, AppName); settings.Recent.Remove(p); return; }
                    if (ConfirmDiscard()) LoadFile(p, null);
                }));
            }
        }

        void AddRecent(string p)
        {
            settings.Recent.RemoveAll(x => string.Equals(x, p, StringComparison.OrdinalIgnoreCase));
            settings.Recent.Insert(0, p);
            while (settings.Recent.Count > 10) settings.Recent.RemoveAt(settings.Recent.Count - 1);
        }

        void OnDragEnter(object s, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void OnDragDrop(object s, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;
            BeginInvoke((Action)(() => { if (ConfirmDiscard()) LoadFile(files[0], null); }));
        }

        void ShowHelp()
        {
            MessageBox.Show(this,
                "■ 基本操作\n" +
                "・セルを選んでそのまま入力、F2 で編集、Enter で確定して下へ、Shift+Enter で確定して上へ\n" +
                "・文字を打ち始めた入力中は、矢印キーで確定して隣のセルへ移動\n" +
                "・F2 / ダブルクリックで開いた編集中は、←→ で文字単位、↑ で先頭、↓ で末尾へ (Shift 併用で選択)\n" +
                "  入力中に F2 を押すと編集中に切り替わります\n" +
                "・列見出し(A,B..)クリックで列選択、行番号クリックで行選択 (Shift で範囲)\n" +
                "・ファイルはウィンドウにドラッグ＆ドロップでも開けます\n" +
                "・表の外側(灰色の行番号)に入力すると自動で行・列が増えます\n\n" +
                "■ ショートカット\n" +
                "Ctrl+N 新規 / Ctrl+O 開く / Ctrl+S 保存 / Ctrl+Shift+S 名前を付けて保存\n" +
                "Ctrl+Z 元に戻す / Ctrl+Y やり直し\n" +
                "Ctrl+X 切り取り / Ctrl+C コピー / Ctrl+V 貼り付け / Del 削除\n" +
                "Ctrl+D 下方向へコピー / Ctrl+R 右方向へコピー\n" +
                "Ctrl+F 検索 / F3 次を検索 / Ctrl+H 置換 / Ctrl+G セルへ移動\n" +
                "Ctrl++ 行の挿入 / Ctrl+- 行の削除\n" +
                "Ctrl+; 文字を大きく / Ctrl+Shift+- 小さく / Ctrl+0 標準 / Ctrl+ホイール\n\n" +
                "■ 表示設定 (表示メニュー)\n" +
                "・文字サイズ、フォント\n" +
                "・セルからはみ出す文字: 折り返す / はみ出す / 切り詰める\n" +
                "表示設定はソフトに保存され、CSV ファイルには影響しません。\n\n" +
                "■ 文字コード\n" +
                "開くときに UTF-8 / Shift_JIS / UTF-16 を自動判別し、同じ文字コードで保存します。\n" +
                "Excel で開く予定なら「UTF-8 (BOM付き)」または「Shift_JIS」がおすすめです。",
                AppName + " の使い方", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void RestoreWindow()
        {
            StartPosition = FormStartPosition.Manual;
            var wa = Screen.PrimaryScreen.WorkingArea;
            int w = S(1100), h = S(700);
            int x = wa.Left + (wa.Width - w) / 2, y = wa.Top + (wa.Height - h) / 2;
            int tx, ty, tw, th;
            if (int.TryParse(settings.Get("X", ""), out tx) && int.TryParse(settings.Get("Y", ""), out ty) &&
                int.TryParse(settings.Get("W", ""), out tw) && int.TryParse(settings.Get("H", ""), out th))
            {
                var r = new Rectangle(tx, ty, tw, th);
                bool visible = Screen.AllScreens.Any(sc => sc.WorkingArea.IntersectsWith(r));
                if (visible && tw > 200 && th > 150) { x = tx; y = ty; w = tw; h = th; }
            }
            Bounds = new Rectangle(x, y, w, h);
            if (settings.Get("Max", "0") == "1") WindowState = FormWindowState.Maximized;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!ConfirmDiscard()) { e.Cancel = true; return; }
            var b = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            settings.Set("X", b.X.ToString()); settings.Set("Y", b.Y.ToString());
            settings.Set("W", b.Width.ToString()); settings.Set("H", b.Height.ToString());
            settings.Set("Max", WindowState == FormWindowState.Maximized ? "1" : "0");
            settings.Save();
            base.OnFormClosing(e);
        }
    }

    // ===================== 検索・置換ダイアログ =====================
    // ===================== グリッド (キー操作のカスタマイズ) =====================
    class CsvGrid : DataGridView
    {
        // Excel と同じ 2 つのモード
        //   入力モード (文字を打ち始めて入力): 矢印キーで確定して隣のセルへ移動
        //   編集モード (F2 / ダブルクリック) : 矢印キーはセル内の文字を移動
        public bool TextNavMode;
        bool pendingTextNav;

        public void BeginEditMode()
        {
            if (CurrentCell == null) return;
            if (IsCurrentCellInEditMode) { TextNavMode = true; return; }
            pendingTextNav = true;
            BeginEdit(false);
            pendingTextNav = false;
        }

        protected override void OnCellBeginEdit(DataGridViewCellCancelEventArgs e)
        {
            TextNavMode = pendingTextNav;
            base.OnCellBeginEdit(e);
        }

        protected override void OnCellEndEdit(DataGridViewCellEventArgs e)
        {
            TextNavMode = false;
            base.OnCellEndEdit(e);
        }

        // 入力中にセル内 (入力欄) をクリックしたら編集モードへ (Excel と同じ)
        TextBox hookedBox;
        protected override void OnEditingControlShowing(DataGridViewEditingControlShowingEventArgs e)
        {
            base.OnEditingControlShowing(e);
            var tb = e.Control as TextBox;
            if (tb != null && tb != hookedBox)
            {
                hookedBox = tb;
                tb.MouseDown += (s, a) => { if (IsCurrentCellInEditMode) TextNavMode = true; };
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left) return;
            var hit = HitTest(e.X, e.Y);
            if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
            {
                if (CurrentCell == null || CurrentCell.RowIndex != hit.RowIndex || CurrentCell.ColumnIndex != hit.ColumnIndex)
                    CurrentCell = this[hit.ColumnIndex, hit.RowIndex];
                BeginEditMode();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // F2: 編集モードで開く (入力モード中なら編集モードに切り替え)
            if (keyData == Keys.F2)
            {
                BeginEditMode();
                return true;
            }
            // 入力モード中の矢印キー: 確定して隣のセルへ
            if (IsCurrentCellInEditMode && !TextNavMode && (keyData & (Keys.Control | Keys.Alt)) == 0)
            {
                int dr = 0, dc = 0;
                switch (keyData & Keys.KeyCode)
                {
                    case Keys.Up: dr = -1; break;
                    case Keys.Down: dr = 1; break;
                    case Keys.Left: dc = -1; break;
                    case Keys.Right: dc = 1; break;
                }
                if (dr != 0 || dc != 0)
                {
                    var c = CurrentCell;
                    if (c == null || !EndEdit()) return true;
                    int r = Math.Max(0, Math.Min(RowCount - 1, c.RowIndex + dr));
                    int col = Math.Max(0, Math.Min(ColumnCount - 1, c.ColumnIndex + dc));
                    CurrentCell = this[col, r];
                    return true;
                }
            }
            // Shift+Enter: 確定して上のセルへ
            if (keyData == (Keys.Shift | Keys.Enter))
            {
                var c = CurrentCell;
                if (c == null) return true;
                if (IsCurrentCellInEditMode && !EndEdit()) return true;
                if (c.RowIndex > 0) CurrentCell = this[c.ColumnIndex, c.RowIndex - 1];
                return true;
            }
            // 入力中の矢印キー等はセル移動させずテキスト内で処理する
            if (IsCurrentCellInEditMode && TextNavMode)
            {
                var tb = EditingControl as TextBox;
                if (tb != null && CsvEditingControl.HandleNavKey(tb, keyData)) return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }

    class CsvCell : DataGridViewTextBoxCell
    {
        public override Type EditType { get { return typeof(CsvEditingControl); } }
    }

    // 入力中は矢印キーでセルを移動しない
    //   ← → : 文字単位でカーソル移動 (端でも隣のセルへは行かない)
    //   ↑   : 先頭へ   ↓ : 末尾へ   (Shift 併用で選択、複数行のときは行移動して端で先頭/末尾へ)
    class CsvEditingControl : DataGridViewTextBoxEditingControl
    {
        public override bool EditingControlWantsInputKey(Keys keyData, bool dataGridViewWantsInputKey)
        {
            Keys key = keyData & Keys.KeyCode;
            if ((keyData & Keys.Alt) == 0)
            {
                switch (key)
                {
                    case Keys.Left: case Keys.Right: case Keys.Up: case Keys.Down:
                    case Keys.Home: case Keys.End:
                        return true;
                }
            }
            if (keyData == (Keys.Shift | Keys.Enter)) return false;
            return base.EditingControlWantsInputKey(keyData, dataGridViewWantsInputKey);
        }

        // 自前で処理したら true (それ以外は TextBox の標準動作に任せる)
        public static bool HandleNavKey(TextBox tb, Keys keyData)
        {
            if ((keyData & Keys.Alt) != 0) return false;
            Keys key = keyData & Keys.KeyCode;
            bool shift = (keyData & Keys.Shift) != 0, ctrl = (keyData & Keys.Control) != 0;
            int selStart = tb.SelectionStart, selEnd = tb.SelectionStart + tb.SelectionLength;
            switch (key)
            {
                case Keys.Up:
                case Keys.Down:
                {
                    if (ctrl) return false;
                    bool up = key == Keys.Up;
                    if (tb.Multiline && tb.TextLength > 0)
                    {
                        int line = tb.GetLineFromCharIndex(up ? selStart : selEnd);
                        bool edge = up ? line == 0 : line >= tb.GetLineFromCharIndex(tb.TextLength);
                        if (!edge) return false;   // 複数行の途中は行移動
                    }
                    if (shift) { if (up) tb.Select(0, selEnd); else tb.Select(selStart, tb.TextLength - selStart); }
                    else tb.Select(up ? 0 : tb.TextLength, 0);
                    tb.ScrollToCaret();
                    return true;
                }
                case Keys.Home:
                case Keys.End:
                {
                    if (tb.Multiline && !ctrl) return false;
                    bool home = key == Keys.Home;
                    if (shift) { if (home) tb.Select(0, selEnd); else tb.Select(selStart, tb.TextLength - selStart); }
                    else tb.Select(home ? 0 : tb.TextLength, 0);
                    tb.ScrollToCaret();
                    return true;
                }
                case Keys.Left:
                    // 先頭でさらに ← を押しても隣のセルへは移動しない
                    if (selStart == 0 && (tb.SelectionLength == 0 || shift)) return true;
                    if (!shift && !ctrl && tb.SelectionLength > 0) { tb.Select(selStart, 0); return true; }
                    return false;
                case Keys.Right:
                    if (selEnd >= tb.TextLength && (tb.SelectionLength == 0 || shift)) return true;
                    if (!shift && !ctrl && tb.SelectionLength > 0) { tb.Select(selEnd, 0); return true; }
                    return false;
            }
            return false;
        }
    }

    class FindForm : Form
    {
        public TextBox FindBox, ReplaceBox;
        CheckBox caseBox, wholeBox, selBox;
        Label replLabel, info;
        Button btnReplace, btnReplaceAll;
        MainForm main;

        public FindForm(MainForm m)
        {
            main = m;
            Text = "検索";
            Font = m.Font;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            KeyPreview = true;
            MaximizeBox = MinimizeBox = false;

            float dpi;
            using (var g = CreateGraphics()) dpi = g.DpiX;
            int tbw = (int)(240 * dpi / 96f);

            var t = new TableLayoutPanel();
            t.ColumnCount = 3; t.RowCount = 5;
            t.AutoSize = true; t.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            t.Padding = new Padding(8);
            t.Dock = DockStyle.Fill;

            FindBox = new TextBox(); FindBox.Width = tbw;
            ReplaceBox = new TextBox(); ReplaceBox.Width = tbw;
            replLabel = Lbl("置換後:");
            caseBox = new CheckBox(); caseBox.Text = "大文字と小文字を区別"; caseBox.AutoSize = true;
            wholeBox = new CheckBox(); wholeBox.Text = "セル全体が一致"; wholeBox.AutoSize = true;
            selBox = new CheckBox(); selBox.Text = "選択範囲のみ置換"; selBox.AutoSize = true;
            info = new Label(); info.AutoSize = true; info.ForeColor = SystemColors.GrayText;

            var bNext = Btn("次を検索", (s, e) => DoFind(false));
            var bPrev = Btn("前を検索", (s, e) => DoFind(true));
            btnReplace = Btn("置換", (s, e) => { main.ReplaceOne(FindBox.Text, ReplaceBox.Text, caseBox.Checked, wholeBox.Checked); info.Text = ""; });
            btnReplaceAll = Btn("すべて置換", (s, e) =>
            {
                int n = main.ReplaceAll(FindBox.Text, ReplaceBox.Text, caseBox.Checked, wholeBox.Checked, selBox.Checked);
                info.Text = n + " 個のセルを置換しました";
            });

            var opts = new FlowLayoutPanel(); opts.AutoSize = true; opts.FlowDirection = FlowDirection.TopDown; opts.WrapContents = false;
            opts.Controls.Add(caseBox); opts.Controls.Add(wholeBox); opts.Controls.Add(selBox);

            t.Controls.Add(Lbl("検索:"), 0, 0); t.Controls.Add(FindBox, 1, 0); t.Controls.Add(bNext, 2, 0);
            t.Controls.Add(replLabel, 0, 1); t.Controls.Add(ReplaceBox, 1, 1); t.Controls.Add(bPrev, 2, 1);
            t.Controls.Add(opts, 1, 2); t.Controls.Add(btnReplace, 2, 2);
            t.Controls.Add(info, 1, 3); t.Controls.Add(btnReplaceAll, 2, 3);
            Controls.Add(t);
            AcceptButton = bNext;
        }

        Label Lbl(string s)
        {
            var l = new Label(); l.Text = s; l.AutoSize = true; l.Anchor = AnchorStyles.Left; l.Margin = new Padding(3, 6, 3, 3);
            return l;
        }

        Button Btn(string s, EventHandler h)
        {
            var b = new Button(); b.Text = s; b.AutoSize = true; b.Dock = DockStyle.Fill; b.Click += h;
            return b;
        }

        void DoFind(bool back)
        {
            info.Text = "";
            if (!main.FindNext(FindBox.Text, back, caseBox.Checked, wholeBox.Checked))
                info.Text = "見つかりませんでした";
        }

        public void SetReplaceMode(bool replace)
        {
            Text = replace ? "検索と置換" : "検索";
            ReplaceBox.Visible = replLabel.Visible = btnReplace.Visible = btnReplaceAll.Visible = selBox.Visible = replace;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Hide(); e.Handled = true; return; }
            if (e.KeyCode == Keys.F3) { DoFind(e.Shift); e.Handled = true; return; }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            base.OnFormClosing(e);
        }
    }

    // ===================== 入力ボックス =====================
    static class InputBox
    {
        public static string Show(IWin32Window owner, string prompt, string title, string def)
        {
            using (var f = new Form())
            {
                f.Text = title;
                f.Font = SystemFonts.MessageBoxFont;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.AutoSize = true; f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                var t = new TableLayoutPanel(); t.AutoSize = true; t.Padding = new Padding(10); t.ColumnCount = 1;
                var l = new Label(); l.Text = prompt; l.AutoSize = true;
                var tb = new TextBox(); tb.Text = def; tb.Dock = DockStyle.Fill;
                var bp = new FlowLayoutPanel(); bp.AutoSize = true; bp.FlowDirection = FlowDirection.RightToLeft; bp.Dock = DockStyle.Fill;
                var ok = new Button(); ok.Text = "OK"; ok.DialogResult = DialogResult.OK; ok.AutoSize = true;
                var cancel = new Button(); cancel.Text = "キャンセル"; cancel.DialogResult = DialogResult.Cancel; cancel.AutoSize = true;
                bp.Controls.Add(cancel); bp.Controls.Add(ok);
                t.Controls.Add(l); t.Controls.Add(tb); t.Controls.Add(bp);
                f.Controls.Add(t);
                f.AcceptButton = ok; f.CancelButton = cancel;
                return f.ShowDialog(owner) == DialogResult.OK ? tb.Text : null;
            }
        }
    }

    // ===================== 設定 (%APPDATA%\SimpleCsvEditor\settings.ini) =====================
    class Settings
    {
        Dictionary<string, string> d = new Dictionary<string, string>();
        public List<string> Recent = new List<string>();
        string Dir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SimpleCsvEditor"); } }
        string FilePath { get { return Path.Combine(Dir, "settings.ini"); } }

        public string Get(string k, string def) { string v; return d.TryGetValue(k, out v) ? v : def; }
        public void Set(string k, string v) { d[k] = v; }

        public void Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return;
                foreach (var line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    int i = line.IndexOf('=');
                    if (i <= 0) continue;
                    string k = line.Substring(0, i), v = line.Substring(i + 1);
                    if (k == "Recent") Recent.Add(v); else d[k] = v;
                }
            }
            catch { }
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var sb = new StringBuilder();
                foreach (var kv in d) sb.Append(kv.Key).Append('=').Append(kv.Value).Append("\r\n");
                foreach (var r in Recent) sb.Append("Recent=").Append(r).Append("\r\n");
                File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
