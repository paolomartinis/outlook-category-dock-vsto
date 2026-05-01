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
        public string Theme { get; set; }

        public override string ToString()
        {
            return Hidden ? Name + " (hidden)" : Name;
        }
    }

    public sealed class CategoryService
    {
        private readonly Outlook.Application application;
        private readonly string hiddenFile;
        private readonly string themesFile;
        private readonly string themeCatalogFile;
        private readonly string settingsFile;
        private readonly HashSet<string> hidden;
        private readonly Dictionary<string, string> themes;
        private readonly HashSet<string> themeCatalog;

        public CategoryService(Outlook.Application application)
        {
            this.application = application;
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "CategoryDockVsto");
            Directory.CreateDirectory(folder);
            hiddenFile = Path.Combine(folder, "hidden-categories.txt");
            themesFile = Path.Combine(folder, "category-themes.txt");
            themeCatalogFile = Path.Combine(folder, "macro-categories.txt");
            settingsFile = Path.Combine(folder, "settings.ini");
            hidden = LoadHidden();
            themes = LoadThemes();
            themeCatalog = LoadThemeCatalog();
            EnsureDefaultThemeCatalog();
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
                    Hidden = isHidden,
                    Theme = GetCategoryTheme(category.Name)
                });
            }

            return result
                .OrderBy(item => item.Theme, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public IReadOnlyList<string> GetThemes()
        {
            return themes.Values
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Concat(themeCatalog)
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public void AddTheme(string name)
        {
            string normalized = NormalizeThemeName(name);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            themeCatalog.Add(normalized);
            SaveThemeCatalog();
        }

        public void RenameTheme(string oldName, string newName)
        {
            string oldNormalized = NormalizeThemeName(oldName);
            string newNormalized = NormalizeThemeName(newName);
            if (string.IsNullOrWhiteSpace(oldNormalized) || string.IsNullOrWhiteSpace(newNormalized))
            {
                return;
            }

            if (string.Equals(oldNormalized, newNormalized, StringComparison.CurrentCultureIgnoreCase))
            {
                AddTheme(newNormalized);
                return;
            }

            themeCatalog.Remove(oldNormalized);
            themeCatalog.Add(newNormalized);

            foreach (string categoryName in themes.Keys.ToList())
            {
                if (string.Equals(themes[categoryName], oldNormalized, StringComparison.CurrentCultureIgnoreCase))
                {
                    themes[categoryName] = newNormalized;
                }
            }

            SaveThemeCatalog();
            SaveThemes();
        }

        public bool DeleteTheme(string name)
        {
            string normalized = NormalizeThemeName(name);
            if (IsDefaultTheme(normalized))
            {
                return false;
            }

            themeCatalog.Remove(normalized);
            foreach (string categoryName in themes.Keys.ToList())
            {
                if (string.Equals(themes[categoryName], normalized, StringComparison.CurrentCultureIgnoreCase))
                {
                    themes.Remove(categoryName);
                }
            }

            SaveThemeCatalog();
            SaveThemes();
            return true;
        }

        public string GetLanguage()
        {
            if (!File.Exists(settingsFile))
            {
                return AppText.English;
            }

            foreach (string line in File.ReadAllLines(settingsFile, Encoding.UTF8))
            {
                string[] parts = line.Split(new[] { '=' }, 2);
                if (parts.Length == 2 && string.Equals(parts[0], "Language", StringComparison.OrdinalIgnoreCase))
                {
                    return string.Equals(parts[1], AppText.Italian, StringComparison.OrdinalIgnoreCase) ? AppText.Italian : AppText.English;
                }
            }

            return AppText.English;
        }

        public void SetLanguage(string language)
        {
            string value = string.Equals(language, AppText.Italian, StringComparison.OrdinalIgnoreCase) ? AppText.Italian : AppText.English;
            File.WriteAllLines(settingsFile, new[] { "Language=" + value }, Encoding.UTF8);
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

        public void AddOrUpdateCategory(string originalName, string name, Outlook.OlCategoryColor color, string theme)
        {
            Outlook.Categories categories = application.Session.Categories;
            Outlook.Category category = FindCategory(originalName);

            if (category == null)
            {
                categories.Add(name, color, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                SetTheme(name, theme);
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

            if (!string.Equals(originalName, name, StringComparison.CurrentCultureIgnoreCase) && themes.ContainsKey(originalName))
            {
                themes.Remove(originalName);
            }

            SetTheme(name, theme);
        }

        public void DeleteCategory(string name)
        {
            application.Session.Categories.Remove(name);
            hidden.Remove(name);
            SaveHidden();
            themes.Remove(name);
            SaveThemes();
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

        private string GetCategoryTheme(string categoryName)
        {
            if (!string.IsNullOrWhiteSpace(categoryName) && themes.TryGetValue(categoryName, out string theme) && !string.IsNullOrWhiteSpace(theme))
            {
                return theme;
            }

            return AppText.Get(GetLanguage(), "NoTheme");
        }

        private void SetTheme(string categoryName, string theme)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
            {
                return;
            }

            string defaultTheme = AppText.Get(GetLanguage(), "NoTheme");
            if (string.IsNullOrWhiteSpace(theme) || string.Equals(theme.Trim(), defaultTheme, StringComparison.CurrentCultureIgnoreCase))
            {
                themes.Remove(categoryName);
            }
            else
            {
                themes[categoryName] = theme.Trim();
            }

            SaveThemes();
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

        private Dictionary<string, string> LoadThemes()
        {
            var result = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            if (!File.Exists(themesFile))
            {
                return result;
            }

            foreach (string line in File.ReadAllLines(themesFile, Encoding.UTF8))
            {
                string[] parts = line.Split(new[] { '\t' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[0]) && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    result[parts[0]] = parts[1];
                }
            }

            return result;
        }

        private void SaveThemes()
        {
            File.WriteAllLines(
                themesFile,
                themes.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase).Select(item => item.Key + "\t" + item.Value),
                Encoding.UTF8);
        }

        private HashSet<string> LoadThemeCatalog()
        {
            if (!File.Exists(themeCatalogFile))
            {
                return new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            }

            return new HashSet<string>(
                File.ReadAllLines(themeCatalogFile, Encoding.UTF8).Select(NormalizeThemeName).Where(item => !string.IsNullOrWhiteSpace(item)),
                StringComparer.CurrentCultureIgnoreCase);
        }

        private void SaveThemeCatalog()
        {
            File.WriteAllLines(themeCatalogFile, themeCatalog.OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase), Encoding.UTF8);
        }

        private void EnsureDefaultThemeCatalog()
        {
            themeCatalog.Add("General");
            themeCatalog.Add("Projects");
            themeCatalog.Add(AppText.Get(GetLanguage(), "NoTheme"));
            SaveThemeCatalog();
        }

        private static string NormalizeThemeName(string name)
        {
            return string.IsNullOrWhiteSpace(name) ? string.Empty : name.Trim();
        }

        private bool IsDefaultTheme(string name)
        {
            return string.Equals(name, "General", StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(name, AppText.Get(GetLanguage(), "NoTheme"), StringComparison.CurrentCultureIgnoreCase);
        }
    }
}

