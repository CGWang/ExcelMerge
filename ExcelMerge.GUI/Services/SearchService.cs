using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelMerge.GUI.Services
{
    public class SearchService
    {
        private const int MaxHistorySize = 10;

        /// <summary>
        /// Updates search history with the given text and saves to settings.
        /// Returns the text if valid, null if empty.
        /// </summary>
        public string UpdateSearchHistory(string text)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(MaxHistorySize).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            return text;
        }

        /// <summary>
        /// Returns the current search history as a list (snapshot).
        /// </summary>
        public List<string> GetSearchHistory()
        {
            return App.Instance.Setting.SearchHistory.ToList();
        }
    }
}
