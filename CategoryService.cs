using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CategoryDockVsto
{
    public sealed class CategoryInfo
    {
        public string Name { get; set; }
        public Outlook.OlCategoryColor Color { get; set; }
        public bool Hidden { get; set; }

        public override string ToString()
        {
            return Hidden ? Name + " (hidden)" : Name;
        }
    }

    public sealed class CategoryService
    {
        private readonly Outlook.Application application;
        private readonly string hiddenFile;
        private readonly HashSet<string> hidden;

        public CategoryService(Outlook.Application application)
        {
            this.application = application;
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "CategoryDockClassic");
            Directory.CreateDirectory(folder);
            hiddenFile = Path.Combine(folder, "hidden-categories.txt");
            hidden = LoadHidden();
        }

        public IReadOnlyList<CategoryInfo> GetCategories(bool includeHidden)
        {
            var result = new List<CategoryInfo>();
            Outlook.Categories categories = application.Session.Categories;

            for (int i = 1; i <= categories.Count; i++)
            {
                Outlook.Category category = categories[i];
                bool isHidden = hidden.Contains(category.Name);
                if (!includeHidden && isHidden)
                {
                    continue;
                }

                result.Add(new CategoryInfo
                {
                    Name = category.Name,
                    Color = category.Color,
                    Hidden = isHidden
                });
            }

            return result.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        public IReadOnlyList<string> GetAppliedCategoriesOnSelectionLead()
        {
            object item = GetLeadSelectedItem();
            if (item == null)
            {
                return new List<string>();
            }

            return SplitCategories(GetCategoriesValue(item));
        }

        public int SelectedCount()
        {
            Outlook.Selection selection = GetSelection();
            return selection == null ? 0 : selection.Count;
        }

        public int ApplyCategoryToSelection(string categoryName)
        {
            return UpdateCategoryOnSelection(categoryName, true);
        }

        public int RemoveCategoryFromSelection(string categoryName)
        {
            return UpdateCategoryOnSelection(categoryName, false);
        }

        public int ClearCategoriesFromSelection()
        {
            Outlook.Selection selection = GetSelection();
            if (selection == null || selection.Count == 0)
            {
                return 0;
            }

            int changed = 0;
            for (int i = 1; i <= selection.Count; i++)
            {
                object item = selection[i];
                try
                {
                    if (!string.IsNullOrWhiteSpace(GetCategoriesValue(item)))
                    {
                        SetCategoriesValue(item, string.Empty);
                        SaveItem(item);
                        changed++;
                    }
                }
                catch
                {
                }
            }

            return changed;
        }

        public void AddOrUpdateCategory(string originalName, string name, Outlook.OlCategoryColor color)
        {
            Outlook.Categories categories = application.Session.Categories;
            Outlook.Category category = FindCategory(originalName);

            if (category == null)
            {
                categories.Add(name, color, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                return;
            }

            if (!string.Equals(category.Name, name, StringComparison.CurrentCultureIgnoreCase))
            {
                category.Name = name;
            }

            if (category.Color != color)
            {
                SetCategoryColor(categories, name, color);
            }

            if (!string.Equals(originalName, name, StringComparison.CurrentCultureIgnoreCase) && hidden.Remove(originalName))
            {
                hidden.Add(name);
                SaveHidden();
            }
        }

        public void DeleteCategory(string name)
        {
            application.Session.Categories.Remove(name);
            hidden.Remove(name);
            SaveHidden();
        }

        public void SetHidden(string name, bool value)
        {
            if (value)
            {
                hidden.Add(name);
            }
            else
            {
                hidden.Remove(name);
            }

            SaveHidden();
        }

        public bool IsHidden(string name)
        {
            return hidden.Contains(name);
        }

        public void SearchByCategories(IEnumerable<string> categories, bool requireAll)
        {
            string query = BuildSearchQuery(categories, requireAll);
            if (string.IsNullOrWhiteSpace(query))
            {
                return;
            }

            Outlook.Explorer explorer = application.ActiveExplorer();
            explorer.Search(query, Outlook.OlSearchScope.olSearchScopeCurrentFolder);
        }

        public string BuildSearchQuery(IEnumerable<string> categories, bool requireAll)
        {
            var parts = categories
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => "category:\"" + item.Replace("\"", "\\\"") + "\"")
                .ToList();

            return string.Join(requireAll ? " AND " : " OR ", parts);
        }

        private int UpdateCategoryOnSelection(string categoryName, bool add)
        {
            Outlook.Selection selection = GetSelection();
            if (selection == null || selection.Count == 0)
            {
                return 0;
            }

            int changed = 0;
            for (int i = 1; i <= selection.Count; i++)
            {
                object item = selection[i];
                if (TryUpdateCategory(item, categoryName, add))
                {
                    changed++;
                }
            }

            return changed;
        }

        private bool TryUpdateCategory(object item, string categoryName, bool add)
        {
            try
            {
                var categories = SplitCategories(GetCategoriesValue(item));
                bool contains = categories.Any(itemName => string.Equals(itemName, categoryName, StringComparison.CurrentCultureIgnoreCase));

                if (add && !contains)
                {
                    categories.Add(categoryName);
                }
                else if (!add && contains)
                {
                    categories = categories
                        .Where(itemName => !string.Equals(itemName, categoryName, StringComparison.CurrentCultureIgnoreCase))
                        .ToList();
                }
                else
                {
                    return false;
                }

                SetCategoriesValue(item, JoinCategories(categories));
                SaveItem(item);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private Outlook.Selection GetSelection()
        {
            Outlook.Explorer explorer = application.ActiveExplorer();
            return explorer == null ? null : explorer.Selection;
        }

        private object GetLeadSelectedItem()
        {
            Outlook.Selection selection = GetSelection();
            if (selection == null || selection.Count == 0)
            {
                return null;
            }

            return selection[1];
        }

        private static string GetCategoriesValue(object item)
        {
            dynamic dynamicItem = item;
            return Convert.ToString(dynamicItem.Categories, CultureInfo.CurrentCulture) ?? string.Empty;
        }

        private static void SetCategoriesValue(object item, string value)
        {
            dynamic dynamicItem = item;
            dynamicItem.Categories = value;
        }

        private static void SaveItem(object item)
        {
            dynamic dynamicItem = item;
            dynamicItem.Save();
        }

        private static List<string> SplitCategories(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new List<string>();
            }

            char listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator.FirstOrDefault();
            if (listSeparator == '\0')
            {
                listSeparator = ',';
            }

            return value
                .Split(new[] { listSeparator, ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length > 0)
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static string JoinCategories(IEnumerable<string> categories)
        {
            string separator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(separator))
            {
                separator = ",";
            }

            return string.Join(separator + " ", categories);
        }

        private Outlook.Category FindCategory(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            Outlook.Categories categories = application.Session.Categories;
            for (int i = 1; i <= categories.Count; i++)
            {
                Outlook.Category category = categories[i];
                if (string.Equals(category.Name, name, StringComparison.CurrentCultureIgnoreCase))
                {
                    return category;
                }
            }

            return null;
        }

        private static void SetCategoryColor(Outlook.Categories categories, string name, Outlook.OlCategoryColor color)
        {
            Outlook.Category category = null;
            for (int i = 1; i <= categories.Count; i++)
            {
                Outlook.Category current = categories[i];
                if (string.Equals(current.Name, name, StringComparison.CurrentCultureIgnoreCase))
                {
                    category = current;
                    break;
                }
            }

            if (category == null)
            {
                categories.Add(name, color, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                return;
            }

            try
            {
                category.Color = color;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                categories.Remove(category.Name);
                categories.Add(name, color, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            }
        }

        private HashSet<string> LoadHidden()
        {
            if (!File.Exists(hiddenFile))
            {
                return new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            }

            return new HashSet<string>(
                File.ReadAllLines(hiddenFile, Encoding.UTF8).Where(item => !string.IsNullOrWhiteSpace(item)),
                StringComparer.CurrentCultureIgnoreCase);
        }

        private void SaveHidden()
        {
            File.WriteAllLines(hiddenFile, hidden.OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase), Encoding.UTF8);
        }
    }
}

