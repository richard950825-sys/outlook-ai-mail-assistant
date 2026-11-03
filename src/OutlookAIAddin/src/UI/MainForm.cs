using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAIAddin
{
    /// <summary>AI 邮件助手独立窗口：阅读摘要/翻译、撰写润色、一句话写邮件、模板管理、设置。</summary>
    internal sealed class MainForm : Form
    {
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color MutedColor = Color.FromArgb(100, 116, 139);
        private static readonly Color AccentColor = Color.FromArgb(37, 99, 235);
        private static readonly Color PanelBack = Color.FromArgb(248, 250, 252);

        private OutlookContext _ctx = new OutlookContext();
        private List<MailTemplate> _templates = new List<MailTemplate>();
        private string _intent = "open";
        private object _pendingItem;
        private bool _pendingAutoRun;
        private bool _busy;
        private bool _loaded;
        private bool _lastPolishHadSelection;
        private string _pendingSelectionText = "";
        private readonly Stopwatch _watch = new Stopwatch();
        private System.Windows.Forms.Timer _busyTimer;
        private string _busyMessage = "";

        // 头部 / 框架
        private Panel _headerPanel;
        private Label _contextTitle;
        private Label _contextInfo;
        private Label _statusLabel;
        private TabControl _tabs;
        private TabPage _tabRead;
        private TabPage _tabPolish;
        private TabPage _tabCompose;
        private TabPage _tabTemplates;
        private TabPage _tabSettings;

        // 阅读助手
        private TextBox _readOutput;
        private ComboBox _translateTarget;
        private Button _btnSummary;
        private Button _btnTranslateRead;

        // 润色
        private Label _selectionInfo;
        private TextBox _polishOutput;
        private Button _btnPolish;
        private Button _btnShorten;
        private Button _btnFormal;
        private Button _btnTranslateEn;
        private Button _btnTranslateZh;
        private Button _btnReplaceSelection;
        private Button _btnReplaceBody;
        private Button _btnInsertCursor;
        private Button _btnCopyPolish;

        // 写邮件
        private ComboBox _templateCombo;
        private TextBox _composeTo;
        private TextBox _instructionBox;
        private TextBox _composeSubject;
        private TextBox _composeBody;
        private Button _btnGenerate;
        private Button _btnNewMail;
        private Button _btnFillDraft;
        private Label _composeStatus;

        // 模板
        private ListBox _templateList;
        private TextBox _tplName;
        private TextBox _tplTo;
        private TextBox _tplDesc;
        private TextBox _tplSignature;
        private TextBox _tplStructure;

        // 设置
        private TextBox _setBaseUrl;
        private TextBox _setModel;
        private TextBox _setApiKey;
        private CheckBox _chkShowKey;
        private ComboBox _setSummaryLang;
        private ComboBox _setTranslateTarget;
        private NumericUpDown _setTemp;
        private Label _setStatus;

        public MainForm()
        {
            BuildUi();
            Load += OnFormLoad;
        }

        /* ==================================================================
         *  界面搭建
         * ================================================================== */

        private void BuildUi()
        {
            Text = "AI 邮件助手";
            Font = UIFont(9f);
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1000, 700);
            MinimumSize = new Size(880, 600);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.White;
            ShowInTaskbar = true;

            BuildHeader();

            _tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(16, 6) };
            _tabRead = BuildReadTab();
            _tabPolish = BuildPolishTab();
            _tabCompose = BuildComposeTab();
            _tabTemplates = BuildTemplatesTab();
            _tabSettings = BuildSettingsTab();
            _tabs.TabPages.AddRange(new[] { _tabRead, _tabPolish, _tabCompose, _tabTemplates, _tabSettings });
            _tabs.SelectedIndexChanged += (s, e) => { };

            var footer = BuildFooter();

            Controls.Add(_tabs);
            Controls.Add(footer);
            Controls.Add(_headerPanel);
        }

        private void BuildHeader()
        {
            _headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 74,
                Padding = new Padding(14, 8, 14, 6),
                BackColor = PanelBack
            };

            _contextTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Font = UIFont(11f, FontStyle.Bold),
                ForeColor = TextColor,
                Text = "正在读取当前邮件…",
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
            _contextInfo = new Label
            {
                Dock = DockStyle.Top,
                Height = 22,
                Font = UIFont(8.5f),
                ForeColor = MutedColor,
                Text = "",
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            var refresh = NewButton("刷新上下文", OnRefreshContext);
            refresh.Dock = DockStyle.Right;
            refresh.Width = 110;
            refresh.Margin = new Padding(0);

            _headerPanel.Controls.Add(_contextInfo);
            _headerPanel.Controls.Add(_contextTitle);
            _headerPanel.Controls.Add(refresh);
        }

        private Panel BuildFooter()
        {
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = Color.White, Padding = new Padding(14, 0, 14, 0) };
            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = "就绪",
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
            var openLog = new LinkLabel
            {
                Dock = DockStyle.Right,
                Width = 110,
                Text = "打开日志文件夹",
                Font = UIFont(8.5f),
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false
            };
            openLog.Click += (s, e) => OpenFolder();
            footer.Controls.Add(_statusLabel);
            footer.Controls.Add(openLog);
            return footer;
        }

        private TabPage BuildReadTab()
        {
            var page = new TabPage("阅读助手") { BackColor = Color.White, Padding = new Padding(0) };
            var layout = NewVerticalLayout(
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.Percent, 100f),
                new RowStyle(SizeType.AutoSize));

            var actions = NewFlowPanel();
            _btnSummary = NewButton("AI 摘要", OnSummaryClick, true);
            _btnTranslateRead = NewButton("AI 翻译", OnTranslateReadClick);
            actions.Controls.Add(_btnSummary);
            actions.Controls.Add(_btnTranslateRead);
            actions.Controls.Add(NewLabel("翻译目标", true));
            _translateTarget = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                Margin = new Padding(0, 4, 0, 0)
            };
            _translateTarget.Items.AddRange(new object[] { "自动（中↔英）", "译成中文", "译成英文" });
            _translateTarget.SelectedIndex = 0;
            actions.Controls.Add(_translateTarget);

            _readOutput = NewOutputBox("选择「AI 摘要」或「AI 翻译」，结果会显示在这里。\r\n\r\n支持三种阅读场景：邮件列表选中 + 右侧阅读窗格、单独打开的邮件窗口，以及右键菜单。");

            var bottom = NewFlowPanel();
            bottom.Controls.Add(NewButton("复制结果", (s, e) => CopyText(_readOutput.Text, "结果")));

            layout.Controls.Add(actions, 0, 0);
            layout.Controls.Add(_readOutput, 0, 1);
            layout.Controls.Add(bottom, 0, 2);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildPolishTab()
        {
            var page = new TabPage("润色（撰写）") { BackColor = Color.White, Padding = new Padding(0) };
            var layout = NewVerticalLayout(
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.Percent, 100f),
                new RowStyle(SizeType.AutoSize));

            _selectionInfo = NewLabel("请打开一封撰写中的邮件", true);
            _selectionInfo.AutoSize = false;
            _selectionInfo.Height = 22;
            _selectionInfo.Dock = DockStyle.Fill;
            _selectionInfo.AutoEllipsis = true;

            var actions = NewFlowPanel();
            _btnPolish = NewButton("润色", (s, e) => SafeRun(() => RunRewriteAsync("polish")), true);
            _btnShorten = NewButton("精简", (s, e) => SafeRun(() => RunRewriteAsync("shorten")));
            _btnFormal = NewButton("正式化", (s, e) => SafeRun(() => RunRewriteAsync("formal")));
            _btnTranslateEn = NewButton("翻译成英文", (s, e) => SafeRun(() => RunTranslateTextAsync("en")));
            _btnTranslateZh = NewButton("翻译成中文", (s, e) => SafeRun(() => RunTranslateTextAsync("zh")));
            actions.Controls.Add(_btnPolish);
            actions.Controls.Add(_btnShorten);
            actions.Controls.Add(_btnFormal);
            actions.Controls.Add(_btnTranslateEn);
            actions.Controls.Add(_btnTranslateZh);

            _polishOutput = NewOutputBox("在打开/新建的邮件正文里选中文字，点「润色」；结果显示在这里，可一键替换回邮件。");

            var bottom = NewFlowPanel();
            _btnReplaceSelection = NewButton("替换选中文字", OnReplaceSelection, true);
            _btnReplaceBody = NewButton("替换整封正文", OnReplaceBody);
            _btnInsertCursor = NewButton("插入到光标处", OnInsertAtCursor);
            _btnCopyPolish = NewButton("复制结果", (s, e) => CopyText(_polishOutput.Text, "结果"));
            bottom.Controls.Add(_btnReplaceSelection);
            bottom.Controls.Add(_btnReplaceBody);
            bottom.Controls.Add(_btnInsertCursor);
            bottom.Controls.Add(_btnCopyPolish);

            layout.Controls.Add(_selectionInfo, 0, 0);
            layout.Controls.Add(actions, 0, 1);
            layout.Controls.Add(_polishOutput, 0, 2);
            layout.Controls.Add(bottom, 0, 3);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildComposeTab()
        {
            var page = new TabPage("写邮件") { BackColor = Color.White, Padding = new Padding(0) };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(10, 8, 10, 8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 86));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var templateRow = NewFlowPanel();
            templateRow.WrapContents = true;
            templateRow.AutoSize = true;
            templateRow.Controls.Add(NewLabel("模板", true));
            _templateCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 320,
                Margin = new Padding(0, 4, 12, 0)
            };
            _templateCombo.SelectedIndexChanged += OnTemplateComboChanged;
            templateRow.Controls.Add(_templateCombo);
            templateRow.Controls.Add(NewLabel("收件人", true));
            _composeTo = new TextBox { Width = 280, Margin = new Padding(0, 4, 0, 0) };
            templateRow.Controls.Add(_composeTo);
            templateRow.Controls.Add(NewButton("管理模板", (s, e) => _tabs.SelectedTab = _tabTemplates));

            var instructionPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            instructionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            instructionPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            instructionPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            instructionPanel.Controls.Add(new Label
            {
                Text = "一句话要求（AI 会按模板结构生成邮件）",
                AutoSize = true,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                Margin = new Padding(0, 0, 0, 4)
            }, 0, 0);
            _instructionBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = UIFont(10f),
                Text = "",
                Margin = new Padding(0)
            };
            instructionPanel.Controls.Add(_instructionBox, 0, 1);

            var generateRow = NewFlowPanel();
            generateRow.WrapContents = true;
            _btnGenerate = NewButton("生成邮件", OnGenerateClick, true);
            _composeStatus = NewLabel("先选模板、写一句话要求，然后点「生成邮件」", true);
            generateRow.Controls.Add(_btnGenerate);
            generateRow.Controls.Add(_composeStatus);

            var subjectRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 2, 0, 2)
            };
            subjectRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            subjectRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            subjectRow.Controls.Add(new Label
            {
                Text = "主题",
                AutoSize = true,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                Margin = new Padding(0, 8, 8, 0)
            }, 0, 0);
            _composeSubject = new TextBox { Dock = DockStyle.Fill, Font = UIFont(10f), Margin = new Padding(0, 4, 0, 4) };
            subjectRow.Controls.Add(_composeSubject, 1, 0);

            _composeBody = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = UIFont(10f),
                AcceptsReturn = true,
                Text = ""
            };

            var bottom = NewFlowPanel();
            bottom.WrapContents = true;
            _btnNewMail = NewButton("新建邮件并填入", OnNewMail, true);
            _btnFillDraft = NewButton("填入当前草稿", OnFillDraft);
            bottom.Controls.Add(_btnNewMail);
            bottom.Controls.Add(_btnFillDraft);
            bottom.Controls.Add(NewButton("复制正文", (s, e) => CopyText(_composeBody.Text, "正文")));

            layout.Controls.Add(templateRow, 0, 0);
            layout.Controls.Add(instructionPanel, 0, 1);
            layout.Controls.Add(generateRow, 0, 2);
            layout.Controls.Add(subjectRow, 0, 3);
            layout.Controls.Add(_composeBody, 0, 4);
            layout.Controls.Add(bottom, 0, 5);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildTemplatesTab()
        {
            var page = new TabPage("模板") { BackColor = Color.White, Padding = new Padding(0) };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(10, 8, 10, 8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var left = new Panel { Dock = DockStyle.Fill };
            _templateList = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
            _templateList.SelectedIndexChanged += OnTemplateSelected;
            var leftButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, AutoSize = false, Padding = new Padding(0, 6, 0, 0) };
            leftButtons.Controls.Add(NewButton("新建", OnTemplateNew));
            leftButtons.Controls.Add(NewButton("删除", OnTemplateDelete));
            left.Controls.Add(_templateList);
            left.Controls.Add(leftButtons);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(10, 0, 0, 0)
            };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _tplName = new TextBox { Margin = new Padding(0, 4, 0, 0) };
            _tplTo = new TextBox { Margin = new Padding(0, 4, 0, 0) };
            _tplDesc = new TextBox { Margin = new Padding(0, 4, 0, 0) };
            _tplSignature = new TextBox { Margin = new Padding(0, 4, 0, 0) };

            right.Controls.Add(FieldRow("模板名称", _tplName), 0, 0);
            right.Controls.Add(FieldRow("默认收件人（可选）", _tplTo), 0, 1);
            right.Controls.Add(FieldRow("说明（可选）", _tplDesc), 0, 2);
            right.Controls.Add(FieldRow("默认签名（可选）", _tplSignature), 0, 3);

            var structurePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            structurePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            structurePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            structurePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            structurePanel.Controls.Add(new Label
            {
                Text = "模板结构 / 风格示例（AI 会照着这个结构和语气来写）",
                AutoSize = true,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                Margin = new Padding(0, 0, 0, 4)
            }, 0, 0);
            _tplStructure = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                Font = UIFont(9.5f),
                Margin = new Padding(0)
            };
            structurePanel.Controls.Add(_tplStructure, 0, 1);
            right.Controls.Add(structurePanel, 0, 4);

            var rightButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            rightButtons.Controls.Add(NewButton("保存模板", OnTemplateSave, true));
            rightButtons.Controls.Add(NewButton("恢复内置模板", OnTemplateRestoreDefaults));
            right.Controls.Add(rightButtons, 0, 5);

            layout.Controls.Add(left, 0, 0);
            layout.Controls.Add(right, 1, 0);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildSettingsTab()
        {
            var page = new TabPage("设置") { BackColor = Color.White, Padding = new Padding(0) };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 11,
                Padding = new Padding(14, 12, 14, 10)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            for (int i = 0; i < 11; i++) { layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); }

            _setBaseUrl = new TextBox { Dock = DockStyle.Fill, Font = UIFont(9.5f) };
            _setModel = new TextBox { Dock = DockStyle.Fill, Font = UIFont(9.5f) };
            _setApiKey = new TextBox { Dock = DockStyle.Fill, Font = UIFont(9.5f), UseSystemPasswordChar = true };
            _chkShowKey = new CheckBox { Text = "显示", AutoSize = true, Margin = new Padding(8, 6, 0, 0) };
            _chkShowKey.CheckedChanged += (s, e) => _setApiKey.UseSystemPasswordChar = !_chkShowKey.Checked;

            _setSummaryLang = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
            _setSummaryLang.Items.AddRange(new object[] { "中文摘要", "跟随邮件原文语言" });
            _setSummaryLang.SelectedIndex = 0;

            _setTranslateTarget = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
            _setTranslateTarget.Items.AddRange(new object[] { "自动（中文→英文 / 其它→中文）", "译成中文", "译成英文" });
            _setTranslateTarget.SelectedIndex = 0;

            _setTemp = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 1,
                DecimalPlaces = 1,
                Increment = 0.1M,
                Width = 80,
                Value = 0.3M
            };

            int row = 0;
            AddSettingRow(layout, row++, "Base URL", _setBaseUrl);
            AddSettingRow(layout, row++, "模型", _setModel);

            var keyPanel = new Panel { Dock = DockStyle.Fill, Height = 28 };
            _setApiKey.Dock = DockStyle.Fill;
            _chkShowKey.Dock = DockStyle.Right;
            _chkShowKey.Width = 60;
            keyPanel.Controls.Add(_setApiKey);
            keyPanel.Controls.Add(_chkShowKey);
            layout.Controls.Add(new Label
            {
                Text = "API Key",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIFont(9f),
                ForeColor = TextColor
            }, 0, 2);
            layout.Controls.Add(keyPanel, 1, 2);
            row = 3;

            AddSettingRow(layout, row++, "摘要语言", _setSummaryLang);
            AddSettingRow(layout, row++, "翻译目标", _setTranslateTarget);
            AddSettingRow(layout, row++, "温度（越低越稳定）", _setTemp);

            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            actions.Controls.Add(NewButton("保存设置", OnSaveSettingsClick, true));
            actions.Controls.Add(NewButton("测试连接", OnTestConnection));
            layout.Controls.Add(actions, 0, row);
            layout.SetColumnSpan(actions, 2);
            row++;

            _setStatus = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 24,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                TextAlign = ContentAlignment.MiddleLeft
            };
            layout.Controls.Add(_setStatus, 0, row);
            layout.SetColumnSpan(_setStatus, 2);
            row++;

            var note = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 70,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                Text = "说明：\r\n" +
                       "1. API Key 使用 Windows DPAPI 加密后保存在本机 %APPDATA%\\OutlookAIAssistant\\settings.json，不会上传到其它地方。\r\n" +
                       "2. 默认接口：DeepSeek（https://api.deepseek.com/v1，模型 deepseek-chat），也可以填任何 OpenAI 兼容接口。\r\n" +
                       "3. 配置目录：" + AppPaths.RootDir
            };
            layout.Controls.Add(note, 0, row);
            layout.SetColumnSpan(note, 2);

            page.Controls.Add(layout);
            return page;
        }

        private static void AddSettingRow(TableLayoutPanel layout, int row, string label, Control control)
        {
            layout.Controls.Add(new Label
            {
                Text = label,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIFont(9f),
                ForeColor = TextColor
            }, 0, row);
            control.Margin = new Padding(0, 4, 0, 4);
            layout.Controls.Add(control, 1, row);
        }

        private Control FieldRow(string label, Control control)
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            row.Controls.Add(new Label
            {
                Text = label,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = MutedColor,
                Font = UIFont(8.5f),
                Margin = new Padding(0, 6, 10, 0)
            }, 0, 0);
            control.Dock = DockStyle.Fill;
            row.Controls.Add(control, 1, 0);
            return row;
        }

        /* ==================================================================
         *  小工具
         * ================================================================== */

        private static Font UIFont(float size)
        {
            return UIFont(size, FontStyle.Regular);
        }

        private static Font UIFont(float size, FontStyle style)
        {
            return new Font("Microsoft YaHei UI", size, style, GraphicsUnit.Point);
        }

        private Button NewButton(string text, EventHandler onClick, bool primary)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                MinimumSize = new Size(0, 30),
                Padding = new Padding(10, 2, 10, 2),
                Margin = new Padding(0, 0, 8, 0),
                FlatStyle = primary ? FlatStyle.Flat : FlatStyle.System
            };
            if (primary)
            {
                button.BackColor = AccentColor;
                button.ForeColor = Color.White;
                button.FlatAppearance.BorderColor = AccentColor;
            }
            if (onClick != null) { button.Click += onClick; }
            return button;
        }

        private Button NewButton(string text, EventHandler onClick)
        {
            return NewButton(text, onClick, false);
        }

        private Label NewLabel(string text, bool muted)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = muted ? MutedColor : TextColor,
                Font = UIFont(muted ? 8.5f : 9f),
                Margin = new Padding(0, 8, 8, 0)
            };
        }

        private static FlowLayoutPanel NewFlowPanel()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Padding = new Padding(0, 2, 0, 2)
            };
        }

        private static TableLayoutPanel NewVerticalLayout(params RowStyle[] rows)
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = rows.Length,
                Padding = new Padding(10, 8, 10, 8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            for (int i = 0; i < rows.Length; i++) { layout.RowStyles.Add(rows[i]); }
            return layout;
        }

        private static TextBox NewOutputBox(string placeholder)
        {
            return new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = true,
                Font = UIFont(10f),
                BackColor = Color.White,
                Text = placeholder,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        /* ==================================================================
         *  上下文
         * ================================================================== */

        private void ApplyContext()
        {
            try
            {
                object item = _pendingItem;
                _pendingItem = null;
                bool preferCompose = _intent == "polish" || _intent == "polish_zh" || _intent == "polish_en";
                _ctx = OutlookContext.Capture(null, item, preferCompose);
            }
            catch (Exception ex)
            {
                Log.Error("ApplyContext", ex);
                _ctx = new OutlookContext { Error = "读取邮件失败：" + ex.Message };
            }
            if (!string.IsNullOrEmpty(_pendingSelectionText) && _ctx.HasItem)
            {
                _ctx.SelectionText = _pendingSelectionText;
            }
            UpdateHeader();
            UpdateActionAvailability();
            Log.Info("当前上下文: hasItem=" + _ctx.HasItem + ", kind=" + _ctx.Kind
                     + ", source=" + _ctx.Source + ", subject=" + (_ctx.Subject ?? "")
                     + ", bodyLength=" + ((_ctx.BodyText ?? "").Length)
                     + ", selectionLength=" + ((_ctx.SelectionText ?? "").Length));
        }

        /// <summary>当前不是撰写上下文时，尝试重新捕获（含阅读窗格内嵌回复）。</summary>
        private bool TryRecaptureComposeContext()
        {
            try
            {
                OutlookContext fresh = OutlookContext.Capture(null, null, true);
                if (fresh != null && fresh.Kind == "compose" && fresh.HasItem)
                {
                    _ctx = fresh;
                    UpdateHeader();
                    UpdateActionAvailability();
                    Log.Info("重新捕获撰写上下文: kind=compose, source=" + fresh.Source
                             + ", selectionLength=" + ((fresh.SelectionText ?? "").Length));
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("TryRecaptureComposeContext: " + ex.Message);
            }
            return false;
        }

        private void UpdateHeader()
        {
            if (_headerPanel == null) { return; }
            if (!_ctx.HasItem)
            {
                _contextTitle.Text = "当前没有可处理的邮件";
                _contextInfo.Text = string.IsNullOrWhiteSpace(_ctx.Error)
                    ? "请选中一封邮件（邮件列表或阅读窗格）、打开一封邮件，或在撰写窗口中再试。"
                    : _ctx.Error;
                _contextInfo.ForeColor = Color.Firebrick;
            }
            else
            {
                _contextTitle.Text = string.IsNullOrWhiteSpace(_ctx.Subject) ? "（无主题）" : _ctx.Subject;
                string kind = _ctx.Kind == "compose" ? "撰写中" : "阅读";
                string body = _ctx.BodyText ?? "";
                string selection = _ctx.SelectionText ?? "";
                string info = _ctx.Source + " · " + kind + " · 正文 " + body.Length + " 字";
                if (!string.IsNullOrEmpty(selection)) { info += " · 选中 " + selection.Length + " 字"; }
                if (_ctx.Kind == "compose")
                {
                    info += selection.Length > 0 ? " · 已选中 " + selection.Length + " 字" : " · 未选中文字";
                    if (!_ctx.WordEditorAvailable) { info += " · 正文编辑器不可用"; }
                }
                _contextInfo.Text = info;
                _contextInfo.ForeColor = MutedColor;
            }

            if (_selectionInfo != null)
            {
                if (_ctx.Kind == "compose")
                {
                    string selection = _ctx.SelectionText ?? "";
                    _selectionInfo.Text = selection.Length > 0
                        ? "已选中 " + selection.Length + " 字：" + FirstLine(selection)
                        : "未选中文字：点「润色/精简/正式化」将处理整封正文";
                }
                else
                {
                    string selection = _ctx.SelectionText ?? "";
                    _selectionInfo.Text = selection.Length > 0
                        ? "（右键菜单）已选中 " + selection.Length + " 字：" + FirstLine(selection)
                        : "提示：在阅读窗格选中文字后右键 → AI 摘要 / AI 翻译；撰写窗口中选中文字右键 → AI 润色。";
                }
            }
        }

        private void UpdateActionAvailability()
        {
            bool read = _ctx.HasItem && _ctx.Kind == "read";
            bool compose = _ctx.HasItem && _ctx.Kind == "compose";

            if (_btnSummary != null) { _btnSummary.Enabled = read && !_busy; }
            if (_btnTranslateRead != null) { _btnTranslateRead.Enabled = read && !_busy; }
            if (_btnPolish != null) { _btnPolish.Enabled = compose && !_busy; }
            if (_btnShorten != null) { _btnShorten.Enabled = compose && !_busy; }
            if (_btnFormal != null) { _btnFormal.Enabled = compose && !_busy; }
            if (_btnTranslateEn != null) { _btnTranslateEn.Enabled = compose && !_busy; }
            if (_btnTranslateZh != null) { _btnTranslateZh.Enabled = compose && !_busy; }
            if (_btnFillDraft != null) { _btnFillDraft.Enabled = compose && !_busy; }
            if (_btnGenerate != null) { _btnGenerate.Enabled = !_busy; }
            if (_btnNewMail != null) { _btnNewMail.Enabled = !_busy; }
            if (_btnReplaceSelection != null) { _btnReplaceSelection.Enabled = compose && _lastPolishHadSelection && !_busy; }
            if (_btnReplaceBody != null) { _btnReplaceBody.Enabled = compose && !_lastPolishHadSelection && !_busy; }
            if (_btnInsertCursor != null) { _btnInsertCursor.Enabled = compose && !_busy; }
        }

        private static string FirstLine(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            int index = text.IndexOf('\n');
            string line = index >= 0 ? text.Substring(0, index) : text;
            if (line.Length > 60) { line = line.Substring(0, 60) + "…"; }
            return line.Replace("\r", "").Trim();
        }

        /* ==================================================================
         *  表单生命周期 / 外部入口
         * ================================================================== */

        private void OnFormLoad(object sender, EventArgs e)
        {
            _loaded = true;
            Log.Info("MainForm 已加载");
            LoadTemplates();
            LoadSettingsUi();
            ApplyContext();
            HandleIntent();
            if (_pendingAutoRun)
            {
                _pendingAutoRun = false;
                RunPendingAutoAction();
            }
        }

        public void Prepare(string intent)
        {
            _intent = NormalizeIntent(intent);
            _pendingItem = null;
            _pendingSelectionText = "";
            _pendingAutoRun = _intent == "summary" || _intent == "translate";
            if (!_loaded) { return; }
            StartFromPending();
        }

        public void PrepareFromItem(object item, string intent)
        {
            _intent = NormalizeIntent(intent);
            _pendingItem = item;
            _pendingSelectionText = "";
            _pendingAutoRun = true;
            if (!_loaded) { return; }
            StartFromPending();
        }

        public void PrepareWithSelection(string intent, string selectionText, object item = null)
        {
            _intent = NormalizeIntent(intent);
            _pendingItem = item;
            _pendingSelectionText = selectionText ?? "";
            _pendingAutoRun = _intent == "summary" || _intent == "translate"
                || _intent == "polish" || _intent == "polish_zh" || _intent == "polish_en";
            if (!_loaded) { return; }
            StartFromPending();
        }

        private void StartFromPending()
        {
            try
            {
                ApplyContext();
                HandleIntent();
                if (_pendingAutoRun)
                {
                    _pendingAutoRun = false;
                    RunPendingAutoAction();
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        private void RunPendingAutoAction()
        {
            if (_intent == "summary") { SafeRun(RunReadSummaryAsync); }
            else if (_intent == "translate") { SafeRun(RunTranslateReadAsync); }
            else if (_intent == "polish") { SafeRun(delegate { return RunRewriteAsync("polish"); }); }
            else if (_intent == "polish_zh") { SafeRun(delegate { return RunRewriteAsync("polish_zh"); }); }
            else if (_intent == "polish_en") { SafeRun(delegate { return RunRewriteAsync("polish_en"); }); }
        }

        private static string NormalizeIntent(string intent)
        {
            if (intent == "summary" || intent == "translate" || intent == "polish"
                || intent == "polish_zh" || intent == "polish_en" || intent == "compose") { return intent; }
            return "open";
        }

        private void HandleIntent()
        {
            switch (_intent)
            {
                case "summary":
                case "translate":
                    if (_tabRead != null) { _tabs.SelectedTab = _tabRead; }
                    break;
                case "polish":
                case "polish_zh":
                case "polish_en":
                    if (_tabPolish != null) { _tabs.SelectedTab = _tabPolish; }
                    break;
                case "compose":
                    if (_tabCompose != null) { _tabs.SelectedTab = _tabCompose; }
                    break;
            }
        }

        /* ==================================================================
         *  通用操作
         * ================================================================== */

        private async void SafeRun(Func<Task> operation)
        {
            try { await operation(); }
            catch (Exception ex) { ShowError(ex); }
        }

        private bool CheckAiConfigured()
        {
            var settings = SettingsStore.Current;
            if (string.IsNullOrWhiteSpace(settings.baseUrl)
                || string.IsNullOrWhiteSpace(settings.model)
                || string.IsNullOrWhiteSpace(SettingsStore.GetApiKey()))
            {
                Warn("还没有配置 AI 接口。请在「设置」里填写 DeepSeek 的 Base URL、模型和 API Key。");
                if (_tabs != null) { _tabs.SelectedTab = _tabSettings; }
                return false;
            }
            return true;
        }

        private bool EnsureItemContext()
        {
            if (_ctx.HasItem) { return true; }
            ApplyContext();
            if (_ctx.HasItem) { return true; }
            Warn(string.IsNullOrWhiteSpace(_ctx.Error) ? "没有可处理的邮件。" : _ctx.Error);
            return false;
        }

        private async Task<AiResult> CallAiAsync(string userPrompt, double temperature, int maxTokens)
        {
            var settings = SettingsStore.Current;
            string apiKey = SettingsStore.GetApiKey();
            return await AiClient.ChatAsync(
                settings.baseUrl, apiKey, settings.model,
                Prompts.SystemPrompt, userPrompt, temperature, maxTokens);
        }

        private void SetBusy(bool busy, string message)
        {
            _busy = busy;
            UseWaitCursor = busy;
            if (busy)
            {
                _busyMessage = string.IsNullOrWhiteSpace(message) ? "AI 正在处理…" : message;
                _watch.Reset();
                _watch.Start();
                UpdateStatus(_busyMessage);
                StartBusyTimer();
                UpdateStatus(_busyMessage);
            }
            else
            {
                _watch.Stop();
                StopBusyTimer();
            }
            UpdateActionAvailability();
        }

        private void StartBusyTimer()
        {
            if (_busyTimer == null)
            {
                _busyTimer = new System.Windows.Forms.Timer { Interval = 1000 };
                _busyTimer.Tick += delegate
                {
                    int seconds = (int)_watch.Elapsed.TotalSeconds;
                    if (seconds < 0) { seconds = 0; }
                    UpdateStatus(_busyMessage + "（已 " + seconds + " 秒）");
                };
            }
            _busyTimer.Start();
        }

        private void StopBusyTimer()
        {
            if (_busyTimer != null) { _busyTimer.Stop(); }
        }

        private string ElapsedText()
        {
            return _watch.Elapsed.TotalSeconds.ToString("0.0") + " 秒";
        }

        private void UpdateStatus(string message)
        {
            RunOnUi(delegate
            {
                if (_statusLabel != null) { _statusLabel.Text = message; }
            });
        }

        private void Info(string message)
        {
            UpdateStatus(message);
        }

        private void Warn(string message)
        {
            RunOnUi(delegate
            {
                try
                {
                    MessageBox.Show(this, message, "AI 邮件助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch
                {
                    // ignore
                }
            });
            UpdateStatus(message);
        }

        private void ShowError(Exception ex)
        {
            Log.Error("操作失败", ex);
            Warn(ex == null ? "未知错误。" : ex.Message);
        }

        private void ShowAiError(string actionLabel, Exception ex, TextBox output)
        {
            Log.Error("AI 操作失败: " + actionLabel, ex);
            string message = ex == null ? "未知错误。" : ex.Message;
            RunOnUi(delegate
            {
                if (output != null)
                {
                    output.Text = "⚠ " + actionLabel + "失败：" + message;
                }
                string shortMessage = message.Length > 60 ? message.Substring(0, 60) + "…" : message;
                if (_statusLabel != null) { _statusLabel.Text = actionLabel + "失败：" + shortMessage; }
            });
        }

        private void RunOnUi(Action action)
        {
            if (action == null) { return; }
            try
            {
                if (IsDisposed) { return; }
                if (InvokeRequired) { BeginInvoke(action); }
                else { action(); }
            }
            catch (Exception ex)
            {
                Log.Error("RunOnUi", ex);
            }
        }

        private void CopyText(string text, string what)
        {
            if (string.IsNullOrEmpty(text))
            {
                Warn("没有可复制的内容。");
                return;
            }
            try
            {
                Clipboard.SetText(text);
                Info("已复制" + what + "到剪贴板。");
            }
            catch (Exception ex)
            {
                Warn("复制失败：" + ex.Message);
            }
        }

        private void OpenFolder()
        {
            try { Process.Start("explorer.exe", "\"" + AppPaths.RootDir + "\""); }
            catch (Exception ex) { Warn("打开文件夹失败：" + ex.Message); }
        }

        /* ==================================================================
         *  阅读助手
         * ================================================================== */

        private async void OnSummaryClick(object sender, EventArgs e)
        {
            await RunReadSummaryAsync();
        }

        private async void OnTranslateReadClick(object sender, EventArgs e)
        {
            await RunTranslateReadAsync();
        }

        private async Task RunReadSummaryAsync()
        {
            if (_busy) { return; }
            try
            {
                if (!CheckAiConfigured()) { return; }
                if (!EnsureItemContext()) { return; }
                if (string.IsNullOrWhiteSpace(_ctx.BodyText)) { Warn("这封邮件没有可读取的正文。"); return; }

                bool useSelection = !string.IsNullOrWhiteSpace(_ctx.SelectionText);
                string sourceText = useSelection ? _ctx.SelectionText : _ctx.BodyText;
                string prompt = Prompts.Summary(_ctx.Subject, _ctx.From, sourceText, SettingsStore.Current.summaryLanguage);
                SetBusy(true, useSelection ? "正在生成摘要（选中片段）…" : "正在生成摘要…");
                try
                {
                    AiResult result = await CallAiAsync(prompt, 0.2, 4000);
                    RunOnUi(delegate
                    {
                        _readOutput.Text = result.Content;
                        _tabs.SelectedTab = _tabRead;
                        UpdateStatus("摘要完成 · " + ElapsedText());
                    });
                }
                finally
                {
                    RunOnUi(delegate { SetBusy(false, ""); });
                }
            }
            catch (Exception ex)
            {
                ShowAiError("摘要", ex, _readOutput);
            }
        }

        private async Task RunTranslateReadAsync()
        {
            if (_busy) { return; }
            try
            {
                if (!CheckAiConfigured()) { return; }
                if (!EnsureItemContext()) { return; }
                if (string.IsNullOrWhiteSpace(_ctx.BodyText)) { Warn("这封邮件没有可读取的正文。"); return; }

                bool useSelection = !string.IsNullOrWhiteSpace(_ctx.SelectionText);
                string sourceText = useSelection ? _ctx.SelectionText : _ctx.BodyText;
                string target = MapTranslateTarget();
                string prompt = Prompts.Translate(sourceText, target);
                SetBusy(true, useSelection ? "正在翻译（选中片段）…" : "正在翻译…");
                try
                {
                    AiResult result = await CallAiAsync(prompt, 0.1, 8000);
                    RunOnUi(delegate
                    {
                        _readOutput.Text = result.Content;
                        _tabs.SelectedTab = _tabRead;
                        UpdateStatus("翻译完成 · " + ElapsedText());
                    });
                }
                finally
                {
                    RunOnUi(delegate { SetBusy(false, ""); });
                }
            }
            catch (Exception ex)
            {
                ShowAiError("翻译", ex, _readOutput);
            }
        }

        private string MapTranslateTarget()
        {
            int index = _translateTarget == null ? 0 : _translateTarget.SelectedIndex;
            if (index == 1) { return "zh"; }
            if (index == 2) { return "en"; }
            string configured = SettingsStore.Current.translateTarget;
            return string.IsNullOrWhiteSpace(configured) ? "auto" : configured;
        }

        /* ==================================================================
         *  润色 / 翻译（撰写）
         * ================================================================== */

        private async Task RunRewriteAsync(string mode)
        {
            if (_busy) { return; }
            try
            {
                if (!CheckAiConfigured()) { return; }
                if (!EnsureItemContext()) { return; }
                if (_ctx.Kind != "compose" && !TryRecaptureComposeContext())
                {
                    Warn("请在打开的邮件撰写窗口，或阅读窗格的内嵌回复中选中文字，再使用润色功能。");
                    if (_tabs != null) { _tabs.SelectedTab = _tabPolish; }
                    return;
                }

                string selected = (_ctx.SelectionText ?? "").Trim();
                bool hasSelection = selected.Length > 0;
                string text = hasSelection ? selected : (_ctx.BodyText ?? "");
                if (text.Trim().Length == 0) { Warn("没有可处理的文字。"); return; }

                _lastPolishHadSelection = hasSelection;
                string prompt;
                string actionName;
                if (mode == "polish_zh" || mode == "polish_en")
                {
                    string targetLanguage = mode == "polish_zh" ? "中文" : "英文";
                    prompt = Prompts.PolishToLanguage(text, targetLanguage, _ctx.Subject, _ctx.From);
                    actionName = "润色成" + targetLanguage;
                }
                else
                {
                    prompt = Prompts.Rewrite(text, mode, _ctx.Subject, _ctx.From);
                    actionName = "润色";
                }
                SetBusy(true, hasSelection ? ("正在" + actionName + "（选中文字）…") : ("正在" + actionName + "…"));
                try
                {
                    AiResult result = await CallAiAsync(prompt, 0.2, 8000);
                    RunOnUi(delegate
                    {
                        _polishOutput.Text = result.Content;
                        _tabs.SelectedTab = _tabPolish;
                        UpdateStatus("处理完成 · " + ElapsedText()
                            + (hasSelection ? " · 可替换选中文字" : " · 将替换整封正文"));
                        UpdateActionAvailability();
                    });
                }
                finally
                {
                    RunOnUi(delegate { SetBusy(false, ""); });
                }
            }
            catch (Exception ex)
            {
                ShowAiError("润色", ex, _polishOutput);
            }
        }

        private async Task RunTranslateTextAsync(string target)
        {
            if (_busy) { return; }
            try
            {
                if (!CheckAiConfigured()) { return; }
                if (!EnsureItemContext()) { return; }
                if (_ctx.Kind != "compose" && !TryRecaptureComposeContext())
                {
                    Warn("翻译选中文字需要在撰写窗口（含阅读窗格内嵌回复）中使用；阅读邮件请用「阅读助手」里的 AI 翻译。");
                    return;
                }

                string selected = (_ctx.SelectionText ?? "").Trim();
                bool hasSelection = selected.Length > 0;
                string text = hasSelection ? selected : (_ctx.BodyText ?? "");
                if (text.Trim().Length == 0) { Warn("没有可翻译的文字。"); return; }

                _lastPolishHadSelection = hasSelection;
                string prompt = Prompts.Translate(text, target);
                SetBusy(true, "正在翻译…");
                try
                {
                    AiResult result = await CallAiAsync(prompt, 0.1, 8000);
                    RunOnUi(delegate
                    {
                        _polishOutput.Text = result.Content;
                        _tabs.SelectedTab = _tabPolish;
                        UpdateStatus("翻译完成 · " + ElapsedText());
                        UpdateActionAvailability();
                    });
                }
                finally
                {
                    RunOnUi(delegate { SetBusy(false, ""); });
                }
            }
            catch (Exception ex)
            {
                ShowAiError("翻译", ex, _polishOutput);
            }
        }

        private void OnReplaceSelection(object sender, EventArgs e)
        {
            if (_busy) { return; }
            if (string.IsNullOrWhiteSpace(_polishOutput.Text)) { Warn("没有可替换的内容。"); return; }
            string result = _ctx.ReplaceSelection(_polishOutput.Text);
            if (result == "ok")
            {
                Info("已替换选中的文字。");
                ApplyContext();
            }
            else
            {
                Warn(result);
            }
        }

        private void OnReplaceBody(object sender, EventArgs e)
        {
            if (_busy) { return; }
            if (string.IsNullOrWhiteSpace(_polishOutput.Text)) { Warn("没有可替换的内容。"); return; }
            if (MessageBox.Show(this, "将用结果替换整封邮件正文（原有正文会被覆盖），继续吗？", "AI 邮件助手",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            {
                return;
            }
            string result = _ctx.ReplaceBodyText(_polishOutput.Text);
            if (result == "ok") { Info("已替换邮件正文。"); ApplyContext(); }
            else { Warn(result); }
        }

        private void OnInsertAtCursor(object sender, EventArgs e)
        {
            if (_busy) { return; }
            if (string.IsNullOrWhiteSpace(_polishOutput.Text)) { Warn("没有可插入的内容。"); return; }
            string result = _ctx.InsertAtCursor(_polishOutput.Text);
            if (result == "ok") { Info("已插入到光标位置。"); ApplyContext(); }
            else { Warn(result); }
        }

        /* ==================================================================
         *  写邮件
         * ================================================================== */

        private async void OnGenerateClick(object sender, EventArgs e)
        {
            await RunComposeAsync();
        }

        private async Task RunComposeAsync()
        {
            if (_busy) { return; }
            try
            {
                if (!CheckAiConfigured()) { return; }

                MailTemplate template = CurrentTemplate();
                string instruction = (_instructionBox.Text ?? "").Trim();
                if (template == null && instruction.Length == 0)
                {
                    Warn("请先选择一个模板，或写一句话说明要写的邮件内容。");
                    return;
                }

                string prompt = Prompts.Compose(template, instruction);
                SetBusy(true, "正在生成邮件…");
                try
                {
                    AiResult result = await CallAiAsync(prompt, 0.6, 6000);
                    var parsed = ComposeResultParser.Parse(result.Content);
                    RunOnUi(delegate
                    {
                        _composeSubject.Text = parsed.Subject ?? "";
                        string body = parsed.Body ?? "";
                        if (template != null && !string.IsNullOrWhiteSpace(template.signature))
                        {
                            if (!body.TrimEnd().EndsWith(template.signature.Trim(), StringComparison.Ordinal))
                            {
                                body = body.TrimEnd() + "\r\n\r\n" + template.signature.Trim();
                            }
                        }
                        _composeBody.Text = body;
                        if (string.IsNullOrWhiteSpace(_composeTo.Text) && template != null && !string.IsNullOrWhiteSpace(template.defaultTo))
                        {
                            _composeTo.Text = template.defaultTo;
                        }
                        _composeStatus.Text = "已生成 · " + ElapsedText() + " · 可编辑后新建邮件";
                        _tabs.SelectedTab = _tabCompose;
                        UpdateStatus("邮件已生成。");
                    });
                }
                finally
                {
                    RunOnUi(delegate { SetBusy(false, ""); });
                }
            }
            catch (Exception ex)
            {
                ShowAiError("写邮件", ex, _composeBody);
            }
        }

        private void OnNewMail(object sender, EventArgs e)
        {
            if (_busy) { return; }
            string subject = (_composeSubject.Text ?? "").Trim();
            string body = _composeBody.Text ?? "";
            if (subject.Length == 0 && body.Trim().Length == 0) { Warn("先生成或填写邮件内容。"); return; }
            string result = OutlookContext.CreateNewMail((_composeTo.Text ?? "").Trim(), subject, body);
            if (result == "ok") { Info("已新建邮件并填入内容，请检查后发送。"); }
            else { Warn(result); }
        }

        private void OnFillDraft(object sender, EventArgs e)
        {
            if (_busy) { return; }
            if (_ctx.Kind != "compose") { Warn("当前没有正在编辑的草稿窗口。"); return; }
            string subject = (_composeSubject.Text ?? "").Trim();
            string body = _composeBody.Text ?? "";
            if (subject.Length == 0 && body.Trim().Length == 0) { Warn("先生成或填写邮件内容。"); return; }
            string result = _ctx.FillCurrent(subject, body, true);
            if (result == "ok") { Info("已填入当前草稿。"); ApplyContext(); }
            else { Warn(result); }
        }

        /* ==================================================================
         *  模板
         * ================================================================== */

        private void LoadTemplates()
        {
            try { _templates = TemplateStore.Load(); }
            catch (Exception ex)
            {
                Log.Error("加载模板失败", ex);
                _templates = new List<MailTemplate>();
            }
            RefreshTemplateUi(null);
        }

        private void RefreshTemplateUi(string selectId)
        {
            if (_templateList != null)
            {
                _templateList.Items.Clear();
                for (int i = 0; i < _templates.Count; i++) { _templateList.Items.Add(_templates[i].name); }
            }
            if (_templateCombo != null)
            {
                _templateCombo.Items.Clear();
                for (int i = 0; i < _templates.Count; i++) { _templateCombo.Items.Add(_templates[i].name); }
            }

            int selectIndex = 0;
            if (!string.IsNullOrEmpty(selectId))
            {
                for (int i = 0; i < _templates.Count; i++)
                {
                    if (_templates[i].id == selectId) { selectIndex = i; break; }
                }
            }
            if (_templates.Count == 0) { ClearTemplateEditor(); return; }

            if (_templateList != null) { _templateList.SelectedIndex = selectIndex; }
            if (_templateCombo != null) { _templateCombo.SelectedIndex = selectIndex; }
            ShowTemplateInEditor(_templates[selectIndex]);
        }

        private MailTemplate CurrentTemplate()
        {
            int index = _templateCombo == null ? -1 : _templateCombo.SelectedIndex;
            if (index >= 0 && index < _templates.Count) { return _templates[index]; }
            return null;
        }

        private void OnTemplateComboChanged(object sender, EventArgs e)
        {
            MailTemplate template = CurrentTemplate();
            if (template != null && string.IsNullOrWhiteSpace(_composeTo.Text) && !string.IsNullOrWhiteSpace(template.defaultTo))
            {
                _composeTo.Text = template.defaultTo;
            }
        }

        private void OnTemplateSelected(object sender, EventArgs e)
        {
            int index = _templateList == null ? -1 : _templateList.SelectedIndex;
            if (index >= 0 && index < _templates.Count)
            {
                ShowTemplateInEditor(_templates[index]);
                if (_templateCombo != null) { _templateCombo.SelectedIndex = index; }
            }
        }

        private void ShowTemplateInEditor(MailTemplate template)
        {
            if (template == null) { return; }
            _tplName.Text = template.name ?? "";
            _tplTo.Text = template.defaultTo ?? "";
            _tplDesc.Text = template.description ?? "";
            _tplSignature.Text = template.signature ?? "";
            _tplStructure.Text = template.structure ?? "";
        }

        private void ClearTemplateEditor()
        {
            if (_tplName == null) { return; }
            _tplName.Text = "";
            _tplTo.Text = "";
            _tplDesc.Text = "";
            _tplSignature.Text = "";
            _tplStructure.Text = "";
        }

        private void OnTemplateNew(object sender, EventArgs e)
        {
            MailTemplate template = TemplateStore.CreateEmpty();
            _templates.Add(template);
            SaveTemplates();
            RefreshTemplateUi(template.id);
        }

        private void OnTemplateDelete(object sender, EventArgs e)
        {
            int index = _templateList == null ? -1 : _templateList.SelectedIndex;
            if (index < 0 || index >= _templates.Count) { return; }
            if (MessageBox.Show(this, "删除模板「" + _templates[index].name + "」？", "AI 邮件助手",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            {
                return;
            }
            _templates.RemoveAt(index);
            SaveTemplates();
            RefreshTemplateUi(null);
        }

        private void OnTemplateSave(object sender, EventArgs e)
        {
            int index = _templateList == null ? -1 : _templateList.SelectedIndex;
            if (index < 0 || index >= _templates.Count) { Warn("请先选择或新建一个模板。"); return; }

            string name = (_tplName.Text ?? "").Trim();
            if (name.Length == 0) { Warn("模板名称不能为空。"); return; }

            MailTemplate template = _templates[index];
            template.name = name;
            template.defaultTo = (_tplTo.Text ?? "").Trim();
            template.description = (_tplDesc.Text ?? "").Trim();
            template.signature = (_tplSignature.Text ?? "").Trim();
            template.structure = _tplStructure.Text ?? "";
            SaveTemplates();
            RefreshTemplateUi(template.id);
            Info("模板已保存。");
        }

        private void OnTemplateRestoreDefaults(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "用内置的 3 个模板替换当前全部模板？", "AI 邮件助手",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            {
                return;
            }
            _templates = TemplateStore.Defaults();
            SaveTemplates();
            RefreshTemplateUi(null);
            Info("已恢复内置模板。");
        }

        private void SaveTemplates()
        {
            try { TemplateStore.Save(_templates); }
            catch (Exception ex)
            {
                Log.Error("保存模板失败", ex);
                Warn("保存模板失败：" + ex.Message);
            }
        }

        /* ==================================================================
         *  设置
         * ================================================================== */

        private void LoadSettingsUi()
        {
            var settings = SettingsStore.Current;
            _setBaseUrl.Text = settings.baseUrl ?? "";
            _setModel.Text = settings.model ?? "";
            _setApiKey.Text = SettingsStore.GetApiKey();
            _chkShowKey.Checked = false;
            _setApiKey.UseSystemPasswordChar = true;
            _setSummaryLang.SelectedIndex = settings.summaryLanguage == "source" ? 1 : 0;
            _setTranslateTarget.SelectedIndex = settings.translateTarget == "zh" ? 1 : (settings.translateTarget == "en" ? 2 : 0);
            decimal temp = (decimal)Math.Min(1.0, Math.Max(0.0, settings.temperature));
            _setTemp.Value = temp;
        }

        private void SaveSettingsFromUi()
        {
            var settings = SettingsStore.Current;
            settings.baseUrl = (_setBaseUrl.Text ?? "").Trim();
            settings.model = (_setModel.Text ?? "").Trim();
            settings.apiKeyEnc = Dpapi.Protect((_setApiKey.Text ?? "").Trim());
            settings.summaryLanguage = _setSummaryLang.SelectedIndex == 1 ? "source" : "zh";
            settings.translateTarget = _setTranslateTarget.SelectedIndex == 1
                ? "zh"
                : (_setTranslateTarget.SelectedIndex == 2 ? "en" : "auto");
            settings.temperature = (double)_setTemp.Value;
            SettingsStore.Save(settings);
        }

        private void OnSaveSettingsClick(object sender, EventArgs e)
        {
            try
            {
                SaveSettingsFromUi();
                _setStatus.ForeColor = Color.SeaGreen;
                _setStatus.Text = "设置已保存。";
                Info("设置已保存。");
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        private async void OnTestConnection(object sender, EventArgs e)
        {
            try
            {
                SaveSettingsFromUi();
                var settings = SettingsStore.Current;
                string apiKey = SettingsStore.GetApiKey();
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    _setStatus.ForeColor = Color.Firebrick;
                    _setStatus.Text = "请先填写 API Key。";
                    return;
                }
                SetBusy(true, "正在测试连接…");
                _setStatus.ForeColor = MutedColor;
                _setStatus.Text = "测试中…";
                string message = await AiClient.TestAsync(settings.baseUrl, apiKey, settings.model);
                RunOnUi(delegate
                {
                    _setStatus.ForeColor = Color.SeaGreen;
                    _setStatus.Text = message;
                    SetBusy(false, message);
                });
            }
            catch (Exception ex)
            {
                Log.Error("测试连接失败", ex);
                RunOnUi(delegate
                {
                    _setStatus.ForeColor = Color.Firebrick;
                    _setStatus.Text = ex.Message;
                    SetBusy(false, "测试连接失败");
                });
            }
        }

        /* ---------------- 头部按钮 ---------------- */

        private void OnRefreshContext(object sender, EventArgs e)
        {
            ApplyContext();
            Info("已刷新当前邮件上下文。");
        }
    }
}
