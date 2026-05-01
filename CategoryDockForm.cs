using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CategoryDockVsto
{
    public sealed class CategoryDockForm : UserControl
    {
        private readonly CategoryService service;
        private readonly TabControl tabs = new TabControl();
        private readonly ListBox assignList = new ListBox();
        private readonly ListBox manageList = new ListBox();
        private readonly ListBox filterList = new ListBox();
        private readonly TextBox queryBox = new TextBox();
        private readonly TextBox nameBox = new TextBox();
        private readonly ComboBox colorBox = new ComboBox();
        private readonly ComboBox themeBox = new ComboBox();
        private readonly ComboBox assignThemeBox = new ComboBox();
        private readonly ComboBox filterThemeBox = new ComboBox();
        private readonly ComboBox languageBox = new ComboBox();
        private readonly ListBox macroThemeList = new ListBox();
        private readonly TextBox macroThemeBox = new TextBox();
        private readonly CheckBox showHiddenBox = new CheckBox();
        private readonly CheckBox requireAllBox = new CheckBox();
        private readonly Label status = new Label();
        private readonly ToolTip toolTip = new ToolTip();
        private string language;
        private bool refreshing;

        public CategoryDockForm(CategoryService service)
        {
            this.service = service;
            Width = 260;
            Height = 620;
            MinimumSize = new Size(120, 320);
            Font = new Font("Segoe UI", 8.5f);
            Margin = new Padding(0);
            language = service.GetLanguage();

            BuildUi();
            RefreshAll();
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                Padding = new Padding(4)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));

            tabs.Dock = DockStyle.Fill;
            tabs.TabPages.Clear();
            tabs.TabPages.Add(BuildAssignTab());
            tabs.TabPages.Add(BuildManageTab());
            tabs.TabPages.Add(BuildFilterTab());
            tabs.TabPages.Add(BuildSettingsTab());

            status.AutoEllipsis = true;
            status.Dock = DockStyle.Fill;
            status.TextAlign = ContentAlignment.MiddleLeft;

            root.Controls.Add(tabs, 0, 0);
            root.Controls.Add(status, 0, 1);
            Controls.Add(root);
        }

        private TabPage BuildAssignTab()
        {
            var tab = new TabPage(T("Assign"));
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            SetupThemeFilter(assignThemeBox);
            assignThemeBox.SelectedIndexChanged += (_, __) =>
            {
                if (!refreshing)
                {
                    RefreshAll();
                }
            };
            layout.Controls.Add(assignThemeBox, 0, 0);

            assignList.Dock = DockStyle.Fill;
            assignList.DrawMode = DrawMode.OwnerDrawFixed;
            assignList.ItemHeight = 22;
            assignList.FormattingEnabled = true;
            assignList.DrawItem += DrawCategoryListItem;
            assignList.DoubleClick += (_, __) => ApplySelected();
            layout.Controls.Add(assignList, 0, 1);

            var apply = Button(T("Apply"), (_, __) => ApplySelected());
            var remove = Button(T("Remove"), (_, __) => RemoveSelected());
            var removeAll = Button(T("RemoveAll"), (_, __) => ClearAllSelected());
            layout.Controls.Add(apply, 0, 2);
            layout.Controls.Add(remove, 0, 3);
            layout.Controls.Add(removeAll, 0, 4);
            tab.Controls.Add(layout);
            return tab;
        }

        private TabPage BuildManageTab()
        {
            var tab = new TabPage(T("Manage"));
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8, ColumnCount = 2, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            showHiddenBox.Text = T("ShowHidden");
            showHiddenBox.Dock = DockStyle.Fill;
            showHiddenBox.AutoSize = false;
            showHiddenBox.CheckedChanged += (_, __) => RefreshAll();
            layout.Controls.Add(showHiddenBox, 0, 0);
            layout.SetColumnSpan(showHiddenBox, 2);

            nameBox.Text = string.Empty;
            nameBox.Dock = DockStyle.Fill;
            nameBox.Margin = new Padding(0, 1, 0, 1);
            nameBox.MinimumSize = new Size(0, 0);
            layout.Controls.Add(nameBox, 0, 1);
            layout.SetColumnSpan(nameBox, 2);

            colorBox.DropDownStyle = ComboBoxStyle.DropDownList;
            colorBox.Dock = DockStyle.Fill;
            colorBox.Margin = new Padding(0, 1, 0, 1);
            colorBox.MinimumSize = new Size(0, 0);
            colorBox.DrawMode = DrawMode.OwnerDrawFixed;
            colorBox.ItemHeight = 20;
            colorBox.DrawItem += DrawColorComboItem;
            colorBox.DataSource = CategoryColorOption.All();
            colorBox.DisplayMember = "Label";
            colorBox.ValueMember = "Color";
            layout.Controls.Add(colorBox, 0, 2);
            layout.SetColumnSpan(colorBox, 2);

            themeBox.DropDownStyle = ComboBoxStyle.DropDown;
            themeBox.Dock = DockStyle.Fill;
            themeBox.Margin = new Padding(0, 1, 0, 1);
            themeBox.MinimumSize = new Size(0, 0);
            layout.Controls.Add(themeBox, 0, 3);
            layout.SetColumnSpan(themeBox, 2);

            manageList.Dock = DockStyle.Fill;
            manageList.Margin = new Padding(0, 2, 0, 2);
            manageList.DrawMode = DrawMode.OwnerDrawFixed;
            manageList.ItemHeight = 22;
            manageList.DrawItem += DrawCategoryListItem;
            manageList.SelectedIndexChanged += (_, __) => LoadSelectedManageCategory();
            layout.Controls.Add(manageList, 0, 4);
            layout.SetColumnSpan(manageList, 2);

            layout.Controls.Add(Button(T("Save"), (_, __) => SaveCategory()), 0, 5);
            layout.Controls.Add(Button(T("Hide"), (_, __) => HideSelected(true)), 1, 5);
            layout.Controls.Add(Button(T("Show"), (_, __) => HideSelected(false)), 0, 6);
            layout.Controls.Add(Button(T("New"), (_, __) => ClearEditor()), 1, 6);
            layout.Controls.Add(Button(T("Delete"), (_, __) => DeleteSelected()), 0, 7);
            layout.Controls.Add(Button(T("Refresh"), (_, __) => RefreshAll()), 1, 7);

            tab.Controls.Add(layout);
            return tab;
        }

        private TabPage BuildFilterTab()
        {
            var tab = new TabPage(T("Filter"));
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 7, ColumnCount = 1, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            queryBox.ReadOnly = true;
            queryBox.Dock = DockStyle.Fill;
            layout.Controls.Add(queryBox, 0, 0);

            SetupThemeFilter(filterThemeBox);
            filterThemeBox.SelectedIndexChanged += (_, __) =>
            {
                if (!refreshing)
                {
                    RefreshAll();
                }
            };
            layout.Controls.Add(filterThemeBox, 0, 1);

            requireAllBox.Text = T("RequireAll");
            requireAllBox.Dock = DockStyle.Fill;
            requireAllBox.AutoSize = false;
            requireAllBox.CheckedChanged += (_, __) => UpdateQuery();
            layout.Controls.Add(requireAllBox, 0, 2);

            filterList.Dock = DockStyle.Fill;
            filterList.DrawMode = DrawMode.OwnerDrawFixed;
            filterList.ItemHeight = 22;
            filterList.FormattingEnabled = true;
            filterList.DrawItem += DrawCategoryListItem;
            filterList.SelectionMode = SelectionMode.MultiExtended;
            filterList.SelectedIndexChanged += (_, __) => UpdateQuery();
            layout.Controls.Add(filterList, 0, 3);

            layout.Controls.Add(Button(T("Search"), (_, __) => RunSearch()), 0, 4);
            layout.Controls.Add(Button(T("Copy"), (_, __) => Clipboard.SetText(queryBox.Text)), 0, 5);
            layout.Controls.Add(Button(T("Refresh"), (_, __) => RefreshAll()), 0, 6);

            tab.Controls.Add(layout);
            return tab;
        }

        private TabPage BuildSettingsTab()
        {
            var tab = new TabPage("\u2699") { ToolTipText = T("Settings") };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 9, ColumnCount = 2, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            var label = new Label { Text = T("Language"), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            languageBox.DropDownStyle = ComboBoxStyle.DropDownList;
            languageBox.Dock = DockStyle.Fill;
            languageBox.Items.Clear();
            languageBox.Items.Add(new LanguageOption("English", AppText.English));
            languageBox.Items.Add(new LanguageOption("Italiano", AppText.Italian));
            languageBox.SelectedIndex = language == AppText.Italian ? 1 : 0;
            languageBox.SelectedIndexChanged += (_, __) => ChangeLanguage();

            layout.Controls.Add(label, 0, 0);
            layout.SetColumnSpan(label, 2);
            layout.Controls.Add(languageBox, 0, 1);
            layout.SetColumnSpan(languageBox, 2);

            var macroLabel = new Label { Text = T("MacroCategories"), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            layout.Controls.Add(macroLabel, 0, 2);
            layout.SetColumnSpan(macroLabel, 2);

            macroThemeList.Dock = DockStyle.Fill;
            macroThemeList.IntegralHeight = false;
            macroThemeList.SelectedIndexChanged += (_, __) => LoadSelectedMacroTheme();
            layout.Controls.Add(macroThemeList, 0, 3);
            layout.SetColumnSpan(macroThemeList, 2);

            macroThemeBox.Dock = DockStyle.Fill;
            macroThemeBox.Margin = new Padding(0, 1, 0, 1);
            layout.Controls.Add(macroThemeBox, 0, 4);
            layout.SetColumnSpan(macroThemeBox, 2);

            layout.Controls.Add(Button(T("New"), (_, __) => ClearMacroThemeEditor()), 0, 5);
            layout.Controls.Add(Button(T("Save"), (_, __) => AddMacroTheme()), 1, 5);
            layout.Controls.Add(Button(T("Rename"), (_, __) => RenameMacroTheme()), 0, 6);
            layout.Controls.Add(Button(T("Delete"), (_, __) => DeleteMacroTheme()), 1, 6);
            layout.Controls.Add(Button(T("Refresh"), (_, __) => RefreshAll()), 0, 7);
            layout.SetColumnSpan(layout.GetControlFromPosition(0, 7), 2);

            tab.Controls.Add(layout);
            return tab;
        }

        private Button Button(string text, EventHandler handler, string tooltip = null)
        {
            var button = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 1, 0, 1),
                MinimumSize = new Size(0, 0),
                AutoSize = false,
                AutoEllipsis = true,
                Font = new Font(Font.FontFamily, 8.0f)
            };
            if (!string.IsNullOrWhiteSpace(tooltip))
            {
                toolTip.SetToolTip(button, tooltip);
            }

            button.Click += handler;
            return button;
        }

        private void RefreshAll()
        {
            if (refreshing)
            {
                return;
            }

            refreshing = true;
            try
            {
                var visibleCategories = service.GetCategories(false);
                var allCategories = service.GetCategories(showHiddenBox.Checked);

                RefreshThemes();
                assignList.Items.Clear();
                AddGroupedCategories(assignList, FilterByTheme(visibleCategories, assignThemeBox.Text));

                manageList.Items.Clear();
                AddGroupedCategories(manageList, allCategories);

                filterList.Items.Clear();
                AddGroupedCategories(filterList, FilterByTheme(visibleCategories, filterThemeBox.Text));

                RefreshMacroThemeList();
                UpdateQuery();
                status.Text = service.SelectedCount() + " " + T("Selected");
            }
            catch (Exception exception)
            {
                Logger.Write(exception);
                status.Text = exception.Message;
            }
            finally
            {
                refreshing = false;
            }
        }

        private void ApplySelected()
        {
            CategoryInfo category = assignList.SelectedItem as CategoryInfo;
            if (category == null)
            {
                return;
            }

            int changed = service.ApplyCategoryToSelection(category.Name);
            status.Text = changed + " " + T("Updated");
            RefreshAll();
        }

        private void RemoveSelected()
        {
            CategoryInfo category = assignList.SelectedItem as CategoryInfo;
            if (category == null)
            {
                return;
            }

            int changed = service.RemoveCategoryFromSelection(category.Name);
            status.Text = changed + " " + T("Updated");
            RefreshAll();
        }

        private void ClearAllSelected()
        {
            int changed = service.ClearCategoriesFromSelection();
            status.Text = changed + " " + T("Cleared");
            RefreshAll();
        }

        private void LoadSelectedManageCategory()
        {
            string name = SelectedManageName();
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            nameBox.Text = name;
            CategoryInfo selected = manageList.SelectedItem as CategoryInfo;
            if (selected != null)
            {
                themeBox.Text = selected.Theme;
                foreach (CategoryColorOption option in colorBox.Items)
                {
                    if (option.Color == selected.Color)
                    {
                        colorBox.SelectedItem = option;
                        break;
                    }
                }
            }
        }

        private void SaveCategory()
        {
            string selected = SelectedManageName();
            string name = nameBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            try
            {
                service.AddOrUpdateCategory(selected, name, ((CategoryColorOption)colorBox.SelectedItem).Color, themeBox.Text);
                status.Text = T("Saved");
                RefreshAll();
            }
            catch (Exception exception)
            {
                Logger.Write(exception);
                MessageBox.Show(exception.Message, "Category Dock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HideSelected(bool hidden)
        {
            string name = SelectedManageName();
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            service.SetHidden(name, hidden);
            RefreshAll();
        }

        private void DeleteSelected()
        {
            string name = SelectedManageName();
            if (string.IsNullOrEmpty(name))
            {
                return;
            }

            if (MessageBox.Show(T("DeleteConfirm") + " \"" + name + "\"?", "Category Dock", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                service.DeleteCategory(name);
                RefreshAll();
            }
        }

        private void ClearEditor()
        {
            manageList.ClearSelected();
            nameBox.Clear();
            themeBox.Text = T("NoTheme");
            if (colorBox.Items.Count > 0)
            {
                colorBox.SelectedIndex = 0;
            }
        }

        private void LoadSelectedMacroTheme()
        {
            if (macroThemeList.SelectedItem == null)
            {
                return;
            }

            macroThemeBox.Text = Convert.ToString(macroThemeList.SelectedItem);
        }

        private void ClearMacroThemeEditor()
        {
            macroThemeList.ClearSelected();
            macroThemeBox.Clear();
            macroThemeBox.Focus();
        }

        private void AddMacroTheme()
        {
            service.AddTheme(macroThemeBox.Text);
            status.Text = T("MacroSaved");
            RefreshAll();
        }

        private void RenameMacroTheme()
        {
            string oldName = Convert.ToString(macroThemeList.SelectedItem);
            if (string.IsNullOrWhiteSpace(oldName))
            {
                AddMacroTheme();
                return;
            }

            service.RenameTheme(oldName, macroThemeBox.Text);
            status.Text = T("MacroSaved");
            RefreshAll();
        }

        private void DeleteMacroTheme()
        {
            string name = Convert.ToString(macroThemeList.SelectedItem);
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (!service.DeleteTheme(name))
            {
                status.Text = T("CannotDeleteGeneral");
                return;
            }

            macroThemeBox.Clear();
            status.Text = T("MacroDeleted");
            RefreshAll();
        }

        private void UpdateQuery()
        {
            queryBox.Text = service.BuildSearchQuery(CheckedFilterNames(), requireAllBox.Checked);
        }

        private void RunSearch()
        {
            service.SearchByCategories(CheckedFilterNames(), requireAllBox.Checked);
        }

        private IEnumerable<string> CheckedFilterNames()
        {
            return filterList.SelectedItems.Cast<object>().OfType<CategoryInfo>().Select(item => item.Name);
        }

        private string SelectedManageName()
        {
            if (manageList.SelectedItem == null)
            {
                return string.Empty;
            }

            CategoryInfo category = manageList.SelectedItem as CategoryInfo;
            return category == null ? string.Empty : category.Name;
        }

        private void DrawCategoryListItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            object item = sender is ComboBox combo ? combo.Items[e.Index] : ((ListBox)sender).Items[e.Index];
            CategoryGroupHeader header = item as CategoryGroupHeader;
            if (header != null)
            {
                using (Brush brush = new SolidBrush(Color.FromArgb(245, 245, 245)))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }

                using (Font headerFont = new Font(e.Font, FontStyle.Bold))
                {
                    TextRenderer.DrawText(
                        e.Graphics,
                        header.Name,
                        headerFont,
                        new Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 2, e.Bounds.Width - 8, e.Bounds.Height - 2),
                        Color.FromArgb(80, 80, 80),
                        TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
                }

                return;
            }

            CategoryInfo category = item as CategoryInfo;
            string text = category == null ? Convert.ToString(item) : category.ToString();
            Color color = category == null ? Color.LightGray : ToDrawingColor(category.Color);
            Rectangle dot = new Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 5, 12, 12);
            using (Brush brush = new SolidBrush(color))
            using (Pen pen = new Pen(Color.FromArgb(110, 110, 110)))
            {
                e.Graphics.FillEllipse(brush, dot);
                e.Graphics.DrawEllipse(pen, dot);
            }

            TextRenderer.DrawText(
                e.Graphics,
                text,
                e.Font,
                new Rectangle(e.Bounds.Left + 22, e.Bounds.Top + 2, e.Bounds.Width - 24, e.Bounds.Height - 2),
                e.ForeColor,
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
            e.DrawFocusRectangle();
        }

        private void DrawColorComboItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            CategoryColorOption option = (CategoryColorOption)colorBox.Items[e.Index];
            Rectangle swatch = new Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 4, 13, 13);
            using (Brush brush = new SolidBrush(ToDrawingColor(option.Color)))
            using (Pen pen = new Pen(Color.FromArgb(110, 110, 110)))
            {
                e.Graphics.FillRectangle(brush, swatch);
                e.Graphics.DrawRectangle(pen, swatch);
            }

            TextRenderer.DrawText(
                e.Graphics,
                option.Label,
                e.Font,
                new Rectangle(e.Bounds.Left + 24, e.Bounds.Top + 1, e.Bounds.Width - 26, e.Bounds.Height - 2),
                e.ForeColor,
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
            e.DrawFocusRectangle();
        }

        private static Color ToDrawingColor(Outlook.OlCategoryColor color)
        {
            switch (color)
            {
                case Outlook.OlCategoryColor.olCategoryColorRed: return Color.FromArgb(232, 72, 85);
                case Outlook.OlCategoryColor.olCategoryColorOrange: return Color.FromArgb(255, 140, 0);
                case Outlook.OlCategoryColor.olCategoryColorPeach: return Color.FromArgb(255, 179, 102);
                case Outlook.OlCategoryColor.olCategoryColorYellow: return Color.FromArgb(252, 225, 0);
                case Outlook.OlCategoryColor.olCategoryColorGreen: return Color.FromArgb(16, 124, 16);
                case Outlook.OlCategoryColor.olCategoryColorTeal: return Color.FromArgb(0, 178, 148);
                case Outlook.OlCategoryColor.olCategoryColorOlive: return Color.FromArgb(186, 216, 10);
                case Outlook.OlCategoryColor.olCategoryColorBlue: return Color.FromArgb(0, 120, 212);
                case Outlook.OlCategoryColor.olCategoryColorPurple: return Color.FromArgb(92, 45, 145);
                case Outlook.OlCategoryColor.olCategoryColorMaroon: return Color.FromArgb(180, 0, 158);
                case Outlook.OlCategoryColor.olCategoryColorSteel: return Color.FromArgb(105, 121, 126);
                case Outlook.OlCategoryColor.olCategoryColorDarkSteel: return Color.FromArgb(59, 58, 57);
                case Outlook.OlCategoryColor.olCategoryColorGray: return Color.FromArgb(122, 117, 116);
                case Outlook.OlCategoryColor.olCategoryColorDarkGray: return Color.FromArgb(59, 58, 57);
                case Outlook.OlCategoryColor.olCategoryColorBlack: return Color.Black;
                case Outlook.OlCategoryColor.olCategoryColorDarkRed: return Color.FromArgb(164, 38, 44);
                case Outlook.OlCategoryColor.olCategoryColorDarkOrange: return Color.FromArgb(202, 80, 16);
                case Outlook.OlCategoryColor.olCategoryColorDarkPeach: return Color.FromArgb(156, 90, 60);
                case Outlook.OlCategoryColor.olCategoryColorDarkYellow: return Color.FromArgb(196, 156, 0);
                case Outlook.OlCategoryColor.olCategoryColorDarkGreen: return Color.FromArgb(73, 130, 5);
                case Outlook.OlCategoryColor.olCategoryColorDarkTeal: return Color.FromArgb(0, 99, 99);
                case Outlook.OlCategoryColor.olCategoryColorDarkOlive: return Color.FromArgb(96, 94, 17);
                case Outlook.OlCategoryColor.olCategoryColorDarkBlue: return Color.FromArgb(0, 69, 120);
                case Outlook.OlCategoryColor.olCategoryColorDarkPurple: return Color.FromArgb(50, 49, 110);
                case Outlook.OlCategoryColor.olCategoryColorDarkMaroon: return Color.FromArgb(116, 39, 116);
                default: return Color.LightGray;
            }
        }

        private sealed class CategoryColorOption
        {
            public Outlook.OlCategoryColor Color { get; set; }
            public string Label { get; set; }

            public static List<CategoryColorOption> All()
            {
                return Enum.GetValues(typeof(Outlook.OlCategoryColor))
                    .Cast<Outlook.OlCategoryColor>()
                    .Where(color => color != Outlook.OlCategoryColor.olCategoryColorNone)
                    .Select(color => new CategoryColorOption { Color = color, Label = CleanColorName(color) })
                    .ToList();
            }

            public override string ToString()
            {
                return Label;
            }
        }

        private static string CleanColorName(Outlook.OlCategoryColor color)
        {
            string name = color.ToString().Replace("olCategoryColor", string.Empty);
            return string.Concat(name.SelectMany((ch, index) =>
                index > 0 && char.IsUpper(ch) ? new[] { ' ', ch } : new[] { ch }));
        }

        private void SetupThemeFilter(ComboBox combo)
        {
            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.Dock = DockStyle.Fill;
            combo.Margin = new Padding(0, 1, 0, 1);
        }

        private void RefreshThemes()
        {
            string assignSelected = string.IsNullOrWhiteSpace(assignThemeBox.Text) ? T("AllThemes") : assignThemeBox.Text;
            string filterSelected = string.IsNullOrWhiteSpace(filterThemeBox.Text) ? T("AllThemes") : filterThemeBox.Text;
            string editTheme = themeBox.Text;
            var filters = new[] { T("AllThemes") }.Concat(service.GetThemes()).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            SetComboItems(assignThemeBox, filters, assignSelected);
            SetComboItems(filterThemeBox, filters, filterSelected);
            SetComboItems(themeBox, service.GetThemes(), string.IsNullOrWhiteSpace(editTheme) ? T("NoTheme") : editTheme);
        }

        private void RefreshMacroThemeList()
        {
            string selected = Convert.ToString(macroThemeList.SelectedItem);
            macroThemeList.BeginUpdate();
            macroThemeList.Items.Clear();
            foreach (string item in service.GetThemes())
            {
                macroThemeList.Items.Add(item);
            }

            object match = macroThemeList.Items.Cast<object>().FirstOrDefault(item => string.Equals(Convert.ToString(item), selected, StringComparison.CurrentCultureIgnoreCase));
            if (match != null)
            {
                macroThemeList.SelectedItem = match;
            }

            macroThemeList.EndUpdate();
        }

        private static void SetComboItems(ComboBox combo, IEnumerable<string> items, string selected)
        {
            combo.BeginUpdate();
            combo.Items.Clear();
            foreach (string item in items)
            {
                combo.Items.Add(item);
            }

            object match = combo.Items.Cast<object>().FirstOrDefault(item => string.Equals(Convert.ToString(item), selected, StringComparison.CurrentCultureIgnoreCase));
            if (match != null)
            {
                combo.SelectedItem = match;
            }
            else if (combo.DropDownStyle == ComboBoxStyle.DropDownList && combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
            else
            {
                combo.Text = selected;
            }

            combo.EndUpdate();
        }

        private IEnumerable<CategoryInfo> FilterByTheme(IEnumerable<CategoryInfo> categories, string theme)
        {
            if (string.IsNullOrWhiteSpace(theme) || string.Equals(theme, T("AllThemes"), StringComparison.CurrentCultureIgnoreCase))
            {
                return categories;
            }

            return categories.Where(item => string.Equals(item.Theme, theme, StringComparison.CurrentCultureIgnoreCase));
        }

        private static void AddGroupedCategories(ListBox list, IEnumerable<CategoryInfo> categories)
        {
            foreach (var group in categories.GroupBy(item => item.Theme).OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
            {
                list.Items.Add(new CategoryGroupHeader(group.Key));
                foreach (CategoryInfo category in group.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
                {
                    list.Items.Add(category);
                }
            }
        }

        private void ChangeLanguage()
        {
            LanguageOption option = languageBox.SelectedItem as LanguageOption;
            if (option == null || option.Code == language)
            {
                return;
            }

            language = option.Code;
            service.SetLanguage(language);
            Controls.Clear();
            BuildUi();
            RefreshAll();
        }

        private string T(string key)
        {
            return AppText.Get(language, key);
        }

        private sealed class CategoryGroupHeader
        {
            public CategoryGroupHeader(string name)
            {
                Name = name;
            }

            public string Name { get; private set; }
        }

        private sealed class LanguageOption
        {
            public LanguageOption(string label, string code)
            {
                Label = label;
                Code = code;
            }

            public string Label { get; private set; }
            public string Code { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }
    }
}

