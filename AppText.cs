using System.Collections.Generic;

namespace CategoryDockVsto
{
    public static class AppText
    {
        public const string English = "en";
        public const string Italian = "it";

        private static readonly Dictionary<string, string> En = new Dictionary<string, string>
        {
            { "Assign", "Assign" },
            { "Manage", "Manage" },
            { "Filter", "Filter" },
            { "Settings", "Settings" },
            { "Apply", "Apply" },
            { "Remove", "Remove" },
            { "RemoveAll", "Remove all" },
            { "ShowHidden", "Show hidden" },
            { "Save", "Save" },
            { "Hide", "Hide" },
            { "Show", "Show" },
            { "New", "New" },
            { "Delete", "Delete" },
            { "Refresh", "Refresh" },
            { "Search", "Search" },
            { "Copy", "Copy" },
            { "RequireAll", "All" },
            { "Language", "Language" },
            { "Theme", "Theme" },
            { "AllThemes", "All themes" },
            { "NoTheme", "General" },
            { "MacroCategories", "Macro categories" },
            { "MacroName", "Macro category name" },
            { "Rename", "Rename" },
            { "CannotDeleteGeneral", "General cannot be deleted" },
            { "MacroDeleted", "Macro category deleted" },
            { "MacroSaved", "Macro category saved" },
            { "Saved", "Category saved" },
            { "Updated", "item(s) updated" },
            { "Cleared", "item(s) cleared" },
            { "Selected", "selected item(s)" },
            { "DeleteConfirm", "Delete category" }
        };

        private static readonly Dictionary<string, string> It = new Dictionary<string, string>
        {
            { "Assign", "Assegna" },
            { "Manage", "Gestisci" },
            { "Filter", "Filtra" },
            { "Settings", "Impostazioni" },
            { "Apply", "Applica" },
            { "Remove", "Rimuovi" },
            { "RemoveAll", "Rimuovi tutte" },
            { "ShowHidden", "Mostra nascoste" },
            { "Save", "Salva" },
            { "Hide", "Nascondi" },
            { "Show", "Mostra" },
            { "New", "Nuova" },
            { "Delete", "Elimina" },
            { "Refresh", "Aggiorna" },
            { "Search", "Cerca" },
            { "Copy", "Copia" },
            { "RequireAll", "Tutte" },
            { "Language", "Lingua" },
            { "Theme", "Macrotema" },
            { "AllThemes", "Tutti i macrotemi" },
            { "NoTheme", "General" },
            { "MacroCategories", "Macrocategorie" },
            { "MacroName", "Nome macrocategoria" },
            { "Rename", "Rinomina" },
            { "CannotDeleteGeneral", "Generale non puo essere cancellata" },
            { "MacroDeleted", "Macrocategoria eliminata" },
            { "MacroSaved", "Macrocategoria salvata" },
            { "Saved", "Categoria salvata" },
            { "Updated", "elemento/i aggiornato/i" },
            { "Cleared", "elemento/i ripulito/i" },
            { "Selected", "elemento/i selezionato/i" },
            { "DeleteConfirm", "Eliminare la categoria" }
        };

        public static string Get(string language, string key)
        {
            Dictionary<string, string> table = language == Italian ? It : En;
            return table.TryGetValue(key, out string value) ? value : key;
        }
    }
}
