using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Excel_Iterate
{
    [Flags]
    internal enum OcrCharacterKinds { None = 0, Chinese = 1, Digits = 2, Letters = 4, Symbols = 8, All = Chinese | Digits | Letters | Symbols }
    internal sealed class OcrPrefixRule
    {
        internal string Id = Guid.NewGuid().ToString("N"), Text = "";
        internal bool Enabled = true;
        internal OcrCharacterKinds Kinds = OcrCharacterKinds.All;
        internal OcrPrefixRule Copy() { return new OcrPrefixRule { Id = Id, Text = Text, Enabled = Enabled, Kinds = Kinds }; }
        internal string Key { get { return OcrFeatures.Normalize(Text).TrimEnd(':'); } }
    }
    internal sealed class OcrPrefixSnapshot
    {
        internal List<OcrPrefixRule> Rules = new List<OcrPrefixRule>();
        internal string Hash = "missing", Warning = "";
        internal bool Writable = true;
    }
    internal static class OcrPrefixStore
    {
        internal static string RootOverride;
        internal static string PathForUser { get { return Path.Combine(RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "OcrPrefixRules.xml"); } }
        private static string Hash(byte[] bytes) { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", ""); }
        internal static void Validate(IList<OcrPrefixRule> rules, bool requireEnabled)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var rule in rules)
            {
                Guid id;
                if (rule == null || !Guid.TryParseExact(rule.Id, "N", out id) || !ids.Add(rule.Id)) throw new InvalidDataException("前缀记录身份无效或重复。");
                if (rule.Key.Length == 0) throw new InvalidDataException("每一行都需要填写字符前缀；不用的空行请删除。");
                if ((rule.Kinds & ~OcrCharacterKinds.All) != 0) throw new InvalidDataException("字符类型无效：" + rule.Text);
                if (rule.Enabled && rule.Kinds == OcrCharacterKinds.None) throw new InvalidDataException("启用的前缀至少需要勾选一种字符类型：" + rule.Text);
                if (rule.Enabled && !keys.Add(rule.Key)) throw new InvalidDataException("启用的前缀重复：" + rule.Text);
            }
            if (requireEnabled && keys.Count == 0) throw new InvalidDataException("请点击「设置前缀」，至少启用一个字符前缀。");
        }
        internal static OcrPrefixSnapshot Read()
        {
            var result = new OcrPrefixSnapshot();
            try
            {
                if (!File.Exists(PathForUser)) return result;
                byte[] bytes = File.ReadAllBytes(PathForUser); result.Hash = Hash(bytes);
                var doc = new XmlDocument { XmlResolver = null };
                using (var memory = new MemoryStream(bytes, false))
                using (var reader = XmlReader.Create(memory, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) doc.Load(reader);
                if (doc.DocumentElement == null || doc.DocumentElement.Name != "OcrPrefixRules") throw new InvalidDataException();
                string version = doc.DocumentElement.GetAttribute("version");
                if (version != "1" && version != "2") throw new InvalidDataException();
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    var row = node as XmlElement;
                    if (row == null || row.Name != "Prefix") throw new InvalidDataException();
                    bool enabled;
                    if (!bool.TryParse(row.GetAttribute("enabled"), out enabled)) throw new InvalidDataException();
                    int kinds = (int)OcrCharacterKinds.All;
                    if (version == "2" && !int.TryParse(row.GetAttribute("kinds"), NumberStyles.None, CultureInfo.InvariantCulture, out kinds)) throw new InvalidDataException();
                    result.Rules.Add(new OcrPrefixRule { Id = row.GetAttribute("id"), Text = row.GetAttribute("text"), Enabled = enabled, Kinds = (OcrCharacterKinds)kinds });
                }
                Validate(result.Rules, false);
            }
            catch { result.Rules.Clear(); result.Writable = false; result.Warning = "前缀规则文件无法读取；原文件已保留，请核对后再保存规则。"; }
            return result;
        }
        internal static OcrPrefixSnapshot Save(OcrPrefixSnapshot previous, IList<OcrPrefixRule> rules)
        {
            Validate(rules, false);
            if (previous == null || !previous.Writable) throw new IOException("原规则文件不可写，未覆盖原文件。");
            string path = PathForUser; Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var lease = new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                string now = File.Exists(path) ? Hash(File.ReadAllBytes(path)) : "missing";
                if (now != previous.Hash) throw new IOException("前缀规则已被另一个窗口修改，请取消并重新打开设置。");
                string temporary = path + "." + Guid.NewGuid().ToString("N").Substring(0, 12) + ".tmp";
                try
                {
                    using (var writer = XmlWriter.Create(temporary, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true }))
                    {
                        writer.WriteStartElement("OcrPrefixRules"); writer.WriteAttributeString("version", "2");
                        foreach (var rule in rules) { writer.WriteStartElement("Prefix"); writer.WriteAttributeString("id", rule.Id); writer.WriteAttributeString("enabled", rule.Enabled.ToString(CultureInfo.InvariantCulture)); writer.WriteAttributeString("text", rule.Text.Trim()); writer.WriteAttributeString("kinds", ((int)rule.Kinds).ToString(CultureInfo.InvariantCulture)); writer.WriteEndElement(); }
                        writer.WriteEndElement();
                    }
                    if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            return Read();
        }
    }
    internal sealed class OcrFieldResult
    {
        internal string RuleId, Prefix, Chosen = "";
        internal bool Manual, FilteredEmpty;
        internal List<string> Candidates = new List<string>();
        internal OcrFieldResult Copy() { return new OcrFieldResult { RuleId = RuleId, Prefix = Prefix, Chosen = Chosen, Manual = Manual, FilteredEmpty = FilteredEmpty, Candidates = new List<string>(Candidates) }; }
    }
    internal static class OcrCombined
    {
        internal static string FilterCharacters(string value, OcrCharacterKinds kinds)
        {
            var result = new StringBuilder();
            for (int at = 0; at < (value ?? "").Length; at++)
            {
                int count = char.IsHighSurrogate(value[at]) && at + 1 < value.Length && char.IsLowSurrogate(value[at + 1]) ? 2 : 1;
                int code = count == 2 ? char.ConvertToUtf32(value[at], value[at + 1]) : value[at];
                if (!char.IsWhiteSpace(value, at))
                {
                    OcrCharacterKinds kind = FileNameEngine.IsChinese(code) ? OcrCharacterKinds.Chinese :
                        code >= '0' && code <= '9' || code >= 0xff10 && code <= 0xff19 ? OcrCharacterKinds.Digits :
                        code >= 'a' && code <= 'z' || code >= 'A' && code <= 'Z' || code >= 0xff21 && code <= 0xff3a || code >= 0xff41 && code <= 0xff5a ? OcrCharacterKinds.Letters : OcrCharacterKinds.Symbols;
                    if ((kinds & kind) != 0) result.Append(value, at, count);
                }
                at += count - 1;
            }
            return result.ToString();
        }
        private static IEnumerable<string> Runs(IEnumerable<OcrSegment> source, IList<OcrPrefixRule> rules)
        {
            var material = source.ToList();
            foreach (var segment in material.Where(s => s.NoLocation).OrderBy(s => s.Page).ThenBy(s => s.Sequence))
                foreach (string line in (segment.Text ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) yield return OcrFeatures.Normalize(line);
            var rows = new List<List<OcrSegment>>();
            foreach (var segment in material.Where(s => !s.NoLocation).OrderBy(s => s.Page).ThenBy(s => s.Top).ThenBy(s => s.Left))
            {
                var row = rows.LastOrDefault(r => r[0].Page == segment.Page && r.Any(p => Math.Min(p.Top + p.Height, segment.Top + segment.Height) - Math.Max(p.Top, segment.Top) >= Math.Min(p.Height, segment.Height) * .5));
                if (row == null) { row = new List<OcrSegment>(); rows.Add(row); } row.Add(segment);
            }
            foreach (var row in rows)
            {
                double gap = Math.Max(1, row.Select(s => s.Height).OrderBy(h => h).ElementAt(row.Count / 2)) * 1.5;
                var run = new List<OcrSegment>();
                foreach (var segment in row.OrderBy(s => s.Left))
                {
                    string next = OcrFeatures.Normalize(segment.Text), prior = OcrFeatures.Normalize(string.Concat(run.Select(s => s.Text)));
                    bool newField = next.Contains(":") && rules.Any(r => prior.Contains(OcrFeatures.Normalize(r.Text)))
                        && !rules.Any(r => next.StartsWith(OcrFeatures.Normalize(r.Text), StringComparison.Ordinal));
                    if (run.Count > 0 && (segment.Left - (run.Last().Left + run.Last().Width) > gap || newField)) { yield return prior; run.Clear(); }
                    run.Add(segment);
                }
                if (run.Count > 0) yield return OcrFeatures.Normalize(string.Concat(run.Select(s => s.Text)));
            }
        }
        private sealed class Mark { internal int Start, ValueStart; internal OcrPrefixRule Rule; }
        internal static List<OcrFieldResult> Extract(IEnumerable<OcrSegment> segments, IList<OcrPrefixRule> rules)
        {
            OcrPrefixStore.Validate(rules, true);
            var enabled = rules.Where(r => r.Enabled).ToList();
            var result = enabled.Select(r => new OcrFieldResult { RuleId = r.Id, Prefix = r.Text }).ToList();
            if (enabled.Count == 1) result[0].Candidates = OcrFeatures.Extract(segments, enabled[0].Text);
            else
            {
                var markers = rules.OrderByDescending(r => OcrFeatures.Normalize(r.Text).Length).ThenByDescending(r => r.Enabled).ToList();
                foreach (string text in Runs(segments, rules))
                {
                    var marks = new List<Mark>();
                    for (int at = 0; at < text.Length; at++)
                    {
                        var rule = markers.FirstOrDefault(r => { string prefix = OcrFeatures.Normalize(r.Text); return at + prefix.Length <= text.Length && string.CompareOrdinal(text, at, prefix, 0, prefix.Length) == 0; });
                        if (rule == null) continue;
                        string prefixText = OcrFeatures.Normalize(rule.Text); int value = at + prefixText.Length;
                        if (!prefixText.EndsWith(":", StringComparison.Ordinal) && value < text.Length && text[value] == ':') value++;
                        marks.Add(new Mark { Start = at, ValueStart = value, Rule = rule }); at = value - 1;
                    }
                    for (int i = 0; i < marks.Count; i++)
                    {
                        var mark = marks[i]; if (!mark.Rule.Enabled) continue;
                        int end = i + 1 < marks.Count ? marks[i + 1].Start : text.Length;
                        string value = text.Substring(mark.ValueStart, end - mark.ValueStart);
                        if (i + 1 < marks.Count) value = value.TrimEnd('、', '，', ',', '；', ';');
                        var field = result.Single(r => r.RuleId == mark.Rule.Id);
                        if (value.Length > 0 && !field.Candidates.Contains(value)) field.Candidates.Add(value);
                    }
                }
            }
            foreach (var field in result)
            {
                var rule = enabled.Single(r => r.Id == field.RuleId);
                int rawCount = field.Candidates.Count;
                field.Candidates = field.Candidates.Select(value => FilterCharacters(value, rule.Kinds)).Where(value => value.Length > 0).Distinct(StringComparer.Ordinal).ToList();
                field.FilteredEmpty = rawCount > 0 && field.Candidates.Count == 0;
                if (field.Candidates.Count == 1) field.Chosen = field.Candidates[0];
            }
            return result;
        }
        internal static void Apply(FileNameEntry entry, IList<OcrPrefixRule> rules)
        {
            if (entry.OcrSegments == null) return;
            var prior = entry.Fields.ToDictionary(f => f.RuleId, StringComparer.Ordinal);
            var fields = Extract(entry.OcrSegments, rules);
            foreach (var field in fields)
            {
                OcrFieldResult old;
                if (prior.TryGetValue(field.RuleId, out old) && old.Manual && OcrFeatures.Normalize(old.Prefix) == OcrFeatures.Normalize(field.Prefix)) { field.Chosen = old.Chosen; field.Manual = true; }
            }
            entry.Fields = fields;
            entry.Candidates = fields.Count == 1 ? fields[0].Candidates.ToList() : new List<string>();
            if (entry.ExplicitBlank) { entry.Feature = ""; entry.Resolved = true; entry.Status = "已明确留空"; return; }
            if (entry.ManualFeature) { entry.Resolved = true; entry.Status = "手工结果，请核对组合规则"; return; }
            Compose(entry);
        }
        internal static void Compose(FileNameEntry entry)
        {
            entry.Resolved = entry.Fields.Count > 0 && entry.Fields.All(f => !string.IsNullOrWhiteSpace(f.Chosen));
            entry.Feature = entry.Resolved ? string.Join("-", entry.Fields.Select(f => f.Chosen)) : "";
            entry.Status = entry.Resolved ? entry.Fields.Any(f => f.Manual) ? "已核对组合" : "已提取" :
                "需核对：" + string.Join("、", entry.Fields.Where(f => string.IsNullOrWhiteSpace(f.Chosen)).Select(f => f.Prefix + (f.Candidates.Count > 1 ? "（多个候选）" : f.FilteredEmpty ? "（筛选后无可用字符）" : "（缺失）")));
        }
        internal static List<OcrPrefixRule> ExampleRules() { return new List<OcrPrefixRule> { new OcrPrefixRule { Text = "磅单号：" }, new OcrPrefixRule { Text = "毛重：" }, new OcrPrefixRule { Text = "销售订单号：" } }; }
    }
    internal sealed class OcrPrefixesForm : Form, IDialogDefaultLayout
    {
        internal readonly DataGridView Grid = new DataGridView { Name = "prefixRules", Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
        private readonly Label status = Ui.Label("", 9, false);
        private readonly bool preview;
        private readonly Func<List<OcrPrefixRule>, string> save;
        internal List<OcrPrefixRule> Result;
        internal OcrPrefixesForm(IList<OcrPrefixRule> rules, Func<List<OcrPrefixRule>, string> save)
        {
            this.save = save; preview = DialogPreviewFactory.IsConstructing; Text = "设置OCR前缀"; DialogSizing.Prepare(this, "file.ocrprefixes", "settings", Ui.Blue); KeyPreview = true;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), RowCount = 4, ColumnCount = 1 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var note = Ui.Label("每行一个前缀，独立勾选保留的字符类型；启用字段按当前顺序用“-”组合，未启用前缀仍作字段边界。字符筛选仅作用于OCR候选，手工值保持原样。", 9, false); note.AutoSize = true; note.MaximumSize = new Size(600, 0); root.Controls.Add(note, 0, 0);
            root.SizeChanged+=delegate {note.MaximumSize=new Size(Math.Max(180,root.ClientSize.Width-root.Padding.Horizontal),0);status.MaximumSize=note.MaximumSize;};
            Ui.StyleGrid(Grid); Grid.RowHeadersVisible = false;
            Grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "enabled", HeaderText = "启用", Width = 70, SortMode = DataGridViewColumnSortMode.NotSortable });
            Grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "prefix", HeaderText = "字符前缀（例如：磅单号：）", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });
            AddKindColumn("chinese", "中文", 58); AddKindColumn("digits", "数字", 58); AddKindColumn("letters", "字母", 58); AddKindColumn("symbols", "其他符号", 90);
            Grid.CurrentCellDirtyStateChanged += delegate { if (Grid.IsCurrentCellDirty && Grid.CurrentCell is DataGridViewCheckBoxCell) Grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            foreach (var rule in rules) Add(rule.Copy()); root.Controls.Add(Grid, 0, 1);
            status.AutoSize = true; status.ForeColor = Ui.Danger; root.Controls.Add(status, 0, 2);
            var actions = Ui.ActionBar(); AddButton(actions, "添加", () => Add(new OcrPrefixRule())); AddButton(actions, "删除", () => { if (Grid.CurrentRow != null) Grid.Rows.Remove(Grid.CurrentRow); });
            AddButton(actions, "上移", () => MoveRule(-1)); AddButton(actions, "下移", () => MoveRule(1));
            var accept = Ui.Button("保存规则", 110); Ui.StylePrimary(accept); accept.Click += delegate { Submit(); }; actions.Controls.Add(accept); AcceptButton = accept;
            var cancel = Ui.Button("取消", 90); cancel.DialogResult = DialogResult.Cancel; actions.Controls.Add(cancel); CancelButton = cancel;
            root.Controls.Add(actions, 0, 3); Controls.Add(root);
        }
        private void AddButton(Control parent, string label, Action action) { var button = Ui.Button(label, 76); button.Click += delegate { if (!preview && !DialogPreviewFactory.IsPreview(this)) action(); }; parent.Controls.Add(button); }
        private void AddKindColumn(string name, string label, int width) { Grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = name, HeaderText = label, Width = width, MinimumWidth = width, AutoSizeMode = DataGridViewAutoSizeColumnMode.None, SortMode = DataGridViewColumnSortMode.NotSortable }); }
        private void Add(OcrPrefixRule rule) { int i = Grid.Rows.Add(rule.Enabled, rule.Text, (rule.Kinds & OcrCharacterKinds.Chinese) != 0, (rule.Kinds & OcrCharacterKinds.Digits) != 0, (rule.Kinds & OcrCharacterKinds.Letters) != 0, (rule.Kinds & OcrCharacterKinds.Symbols) != 0); Grid.Rows[i].Tag = rule; Grid.CurrentCell = Grid.Rows[i].Cells[1]; }
        public Size MeasureDefaultClient() { return new Size(680,440); }
        protected override bool ProcessCmdKey(ref Message message,Keys keys)
        {if(keys==Keys.Escape){DialogResult=DialogResult.Cancel;Close();return true;}return base.ProcessCmdKey(ref message,keys);}
        private void MoveRule(int direction)
        {
            Grid.EndEdit(); if (Grid.CurrentRow == null) return; int from = Grid.CurrentRow.Index, to = from + direction;
            if (to < 0 || to >= Grid.Rows.Count) return;
            var row = Grid.Rows[from]; Grid.Rows.RemoveAt(from); Grid.Rows.Insert(to, row); Grid.CurrentCell = row.Cells[1];
        }
        internal void Submit()
        {
            if (preview || DialogPreviewFactory.IsPreview(this)) return;
            try
            {
                Grid.EndEdit(); var rules = new List<OcrPrefixRule>();
                foreach (DataGridViewRow row in Grid.Rows)
                {
                    var rule = ((OcrPrefixRule)row.Tag).Copy(); rule.Text = Convert.ToString(row.Cells[1].Value).Trim(); rule.Enabled = Convert.ToBoolean(row.Cells[0].Value ?? false); rule.Kinds = OcrCharacterKinds.None;
                    var kinds = new[] { OcrCharacterKinds.Chinese, OcrCharacterKinds.Digits, OcrCharacterKinds.Letters, OcrCharacterKinds.Symbols };
                    for (int i = 0; i < kinds.Length; i++) if (Convert.ToBoolean(row.Cells[i + 2].Value ?? false)) rule.Kinds |= kinds[i];
                    rules.Add(rule);
                }
                OcrPrefixStore.Validate(rules, false); string error = save == null ? "无法保存规则。" : save(rules);
                if (!string.IsNullOrEmpty(error)) { status.Text = error; return; }
                Result = rules; DialogResult = DialogResult.OK; Close();
            }
            catch (Exception error) { status.Text = error.Message; }
        }
    }
    internal sealed class OcrFieldsForm : Form, IDialogDefaultLayout
    {
        internal readonly DataGridView Grid = new DataGridView { Name = "ocrFields", Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
        private readonly Label status = Ui.Label("", 9, false);
        private readonly bool preview;
        internal List<OcrFieldResult> Result;
        internal bool Blank;
        public Size MeasureDefaultClient() { return new Size(820,500); }
        internal OcrFieldsForm(IList<OcrFieldResult> fields)
        {
            preview = DialogPreviewFactory.IsConstructing; Text = "核对OCR字段"; DialogSizing.Prepare(this, "file.ocrfields", "settings", Ui.Blue); KeyPreview = true;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), RowCount = 4, ColumnCount = 1 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var note = Ui.Label("选择候选，或在“使用值”补填。全部字段有值后组合；也可将该文件整行留空。", 9, false); note.AutoSize = true; root.Controls.Add(note, 0, 0);
            root.SizeChanged+=delegate {note.MaximumSize=new Size(Math.Max(180,root.ClientSize.Width-root.Padding.Horizontal),0);status.MaximumSize=note.MaximumSize;};
            Ui.StyleGrid(Grid); Grid.RowHeadersVisible = false;
            Grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "prefix", HeaderText = "前缀", ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            Grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "candidate", HeaderText = "识别候选", SortMode = DataGridViewColumnSortMode.NotSortable });
            Grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "chosen", HeaderText = "使用值（可填写）", SortMode = DataGridViewColumnSortMode.NotSortable });
            foreach (var field in fields)
            {
                int i = Grid.Rows.Add(); var row = Grid.Rows[i]; row.Tag = field.Copy(); row.Cells[0].Value = field.Prefix;
                var choices = (DataGridViewComboBoxCell)row.Cells[1]; choices.Items.Add(""); foreach (string value in field.Candidates) choices.Items.Add(value);
                row.Cells[1].Value = field.Candidates.Contains(field.Chosen) ? field.Chosen : ""; row.Cells[2].Value = field.Chosen;
            }
            Grid.CurrentCellDirtyStateChanged += delegate { if (Grid.IsCurrentCellDirty) Grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            Grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && e.ColumnIndex == 1) { ((OcrFieldResult)Grid.Rows[e.RowIndex].Tag).Manual = true; Grid.Rows[e.RowIndex].Cells[2].Value = Convert.ToString(Grid.Rows[e.RowIndex].Cells[1].Value); } };
            Grid.CellEndEdit += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && (e.ColumnIndex == 1 || e.ColumnIndex == 2)) ((OcrFieldResult)Grid.Rows[e.RowIndex].Tag).Manual = true; };
            Grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; status.Text = "请选择候选，或在使用值列补填。"; };
            root.Controls.Add(Grid, 0, 1); status.AutoSize = true; status.ForeColor = Ui.Danger; root.Controls.Add(status, 0, 2);
            var filtered = fields.Where(field => field.FilteredEmpty && string.IsNullOrWhiteSpace(field.Chosen)).Select(field => field.Prefix).ToList();
            if (filtered.Count > 0) status.Text = "筛选后无可用字符，请补填或整行留空：" + string.Join("、", filtered);
            var actions = Ui.ActionBar(); var accept = Ui.Button("确定组合", 112); Ui.StylePrimary(accept); accept.Click += delegate { Submit(); }; actions.Controls.Add(accept); AcceptButton = accept;
            var blank = Ui.Button("整行留空", 112); blank.Click += delegate { if (preview || DialogPreviewFactory.IsPreview(this)) return; Blank = true; DialogResult = DialogResult.OK; Close(); }; actions.Controls.Add(blank);
            var cancel = Ui.Button("取消", 90); cancel.DialogResult = DialogResult.Cancel; actions.Controls.Add(cancel); CancelButton = cancel;
            root.Controls.Add(actions, 0, 3); Controls.Add(root);
        }
        internal void Submit()
        {
            if (preview || DialogPreviewFactory.IsPreview(this)) return;
            Grid.EndEdit(); var result = new List<OcrFieldResult>();
            foreach (DataGridViewRow row in Grid.Rows)
            {
                var field = ((OcrFieldResult)row.Tag).Copy(); string chosen = Convert.ToString(row.Cells[2].Value).Trim();
                if (chosen.Length == 0) { status.Text = "请补填或选择：" + field.Prefix; return; }
                field.Manual = field.Manual || chosen != field.Chosen || field.Candidates.Count != 1; field.Chosen = chosen; result.Add(field);
            }
            Result = result; DialogResult = DialogResult.OK; Close();
        }
        protected override bool ProcessCmdKey(ref Message message,Keys keys)
        {if(keys==Keys.Escape){DialogResult=DialogResult.Cancel;Close();return true;}return base.ProcessCmdKey(ref message,keys);}
    }
}
