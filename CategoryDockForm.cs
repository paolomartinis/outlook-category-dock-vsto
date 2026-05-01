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
        private readonly CheckBox showHiddenBox = new CheckBox();
        private readonly CheckBox requireAllBox = new CheckBox();
        private readonly Label status = new Label();
        private readonly ToolTip toolTip = new ToolTip();

        public CategoryDockForm(CategoryService service)
        {
            this.service = service;
            Width = 260;
            Height = 620;
            MinimumSize = new Size(120, 320);
            Font = new Font("Segoe UI", 8.5f);
            Margin = new Padding(0);

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
            tabs.TabPages.Add(BuildAssignTab());
            tabs.TabPages.Add(BuildManageTab());
            tabs.TabPages.Add(BuildFilterTab());

            status.AutoEllipsis = true;
            status.Dock = DockStyle.Fill;
            status.TextAlign = ContentAlignment.MiddleLeft;

            root.Controls.Add(tabs, 0, 0);
            root.Controls.Add(status, 0, 1);
            Controls.Add(root);
        }

        private TabPage BuildAssignTab()
        {
            var tab = new TabPage("Assegna");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            assignList.Dock = DockStyle.Fill;
            assignList.DrawMode = DrawMode.OwnerDrawFixed;
            assignList.ItemHeight = 22;
            assignList.FormattingEnabled = true;
            assignList.DrawItem += DrawCategoryListItem;
            assignList.DoubleClick += (_, __) => ApplySelected();
            layout.Controls.Add(assignList, 0, 0);

            var apply = Button("Applica", (_, __) => ApplySelected());
            var remove = Button("Rimuovi", (_, __) => RemoveSelected());
            var removeAll = Button("Rimuovi tutte", (_, __) => ClearAllSelected());
            layout.Controls.Add(apply, 0, 1);
            layout.Controls.Add(remove, 0, 2);
            layout.Controls.Add(removeAll, 0, 3);
            tab.Controls.Add(layout);
            return tab;
        }

        private TabPage BuildManageTab()
        {
            var tab = new TabPage("Gestisci");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 7, ColumnCount = 2, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            showHiddenBox.Text = "Mostra nascoste";
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

            manageList.Dock = DockStyle.Fill;
            manageList.Margin = new Padding(0, 2, 0, 2);
            manageList.DrawMode = DrawMode.OwnerDrawFixed;
            manageList.ItemHeight = 22;
            manageList.DrawItem += DrawCategoryListItem;
            manageList.SelectedIndexChanged += (_, __) => LoadSelectedManageCategory();
            layout.Controls.Add(manageList, 0, 3);
            layout.SetColumnSpan(manageList, 2);

            layout.Controls.Add(Button("Salva", (_, __) => SaveCategory()), 0, 4);
            layout.Controls.Add(Button("Nascondi", (_, __) => HideSelected(true)), 1, 4);
            layout.Controls.Add(Button("Mostra", (_, __) => HideSelected(false)), 0, 5);
            layout.Controls.Add(Button("Nuova", (_, __) => ClearEditor()), 1, 5);
            layout.Controls.Add(Button("Elimina", (_, __) => DeleteSelected()), 0, 6);
            layout.Controls.Add(Button("Aggiorna", (_, __) => RefreshAll()), 1, 6);

            tab.Controls.Add(layout);
            return tab;
        }

        private TabPage BuildFilterTab()
        {
            var tab = new TabPage("Filtra");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 6, ColumnCount = 1, Padding = new Padding(4) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            queryBox.ReadOnly = true;
            queryBox.Dock = DockStyle.Fill;
            layout.Controls.Add(queryBox, 0, 0);

            requireAllBox.Text = "Tutte";
            requireAllBox.Dock = DockStyle.Fill;
            requireAllBox.AutoSize = false;
            requireAllBox.CheckedChanged += (_, __) => UpdateQuery();
            layout.Controls.Add(requireAllBox, 0, 1);

            filterList.Dock = DockStyle.Fill;
            filterList.DrawMode = DrawMode.OwnerDrawFixed;
            filterList.ItemHeight = 22;
            filterList.FormattingEnabled = true;
            filterList.DrawItem += DrawCategoryListItem;
            filterList.SelectionMode = SelectionMode.MultiExtended;
            filterList.SelectedIndexChanged += (_, __) => UpdateQuery();
            layout.Controls.Add(filterList, 0, 2);

            layout.Controls.Add(Button("Cerca", (_, __) => RunSearch()), 0, 3);
            layout.Controls.Add(Button("Copia", (_, __) => Clipboard.SetText(queryBox.Text)), 0, 4);
            layout.Controls.Add(Button("Aggiorna", (_, __) => RefreshAll()), 0, 5);

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
            var visibleCategories = service.GetCategories(false);
            var allCategories = service.GetCategories(showHiddenBox.Checked);
            var applied = new HashSet<string>(service.GetAppliedCategoriesOnSelectionLead(), StringComparer.CurrentCultureIgnoreCase);

            assignList.Items.Clear();
            foreach (var category in visibleCategories)
            {
                assignList.Items.Add(category);
            }

            manageList.Items.Clear();
            foreach (var category in allCategories)
            {
                manageList.Items.Add(category);
            }

            filterList.Items.Clear();
            foreach (var category in visibleCategories)
            {
                filterList.Items.Add(category);
            }

            UpdateQuery();
            status.Text = service.SelectedCount() + " selected item(s)";
        }

        private void ApplySelected()
        {
            CategoryInfo category = assignList.SelectedItem as CategoryInfo;
            if (category == null)
            {
                return;
            }

            int changed = service.ApplyCategoryToSelection(category.Name);
            status.Text = changed + " item(s) updated";
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
            status.Text = changed + " item(s) updated";
            RefreshAll();
        }

        private void ClearAllSelected()
        {
            int changed = service.ClearCategoriesFromSelection();
            status.Text = changed + " item(s) cleared";
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
                service.AddOrUpdateCategory(selected, name, ((CategoryColorOption)colorBox.SelectedItem).Color);
                status.Text = "Category saved";
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

            if (MessageBox.Show("Delete category \"" + name + "\"?", "Category Dock", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                service.DeleteCategory(name);
                RefreshAll();
            }
        }

        private void ClearEditor()
        {
            manageList.ClearSelected();
            nameBox.Clear();
            if (colorBox.Items.Count > 0)
            {
                colorBox.SelectedIndex = 0;
            }
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
            return filterList.SelectedItems.Cast<CategoryInfo>().Select(item => item.Name);
        }

        private string SelectedManageName()
        {
            if (manageList.SelectedItem == null)
            {
                return string.Empty;
            }

            return ((CategoryInfo)manageList.SelectedItem).Name;
        }

        private void DrawCategoryListItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            object item = sender is ComboBox combo ? combo.Items[e.Index] : ((ListBox)sender).Items[e.Index];
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
    }
}

