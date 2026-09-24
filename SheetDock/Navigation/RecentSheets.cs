using System.Collections.Generic;

namespace SheetDock
{
    /// <summary>
    /// History of visited sheets, supporting back/forward navigation.
    /// </summary>
    public class RecentSheetStack
    {
        private readonly List<string> _history = new List<string>();
        private int _index = -1;

        public int Count => _history.Count;
        public bool CanGoBack => _index > 0;
        public bool CanGoForward => _index >= 0 && _index < _history.Count - 1;

        public string Current => (_index >= 0 && _index < _history.Count)
            ? _history[_index]
            : null;

        public void Push(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) return;

            if (_index < _history.Count - 1)
                _history.RemoveRange(_index + 1, _history.Count - _index - 1);

            if (_history.Count > 0 && _history[_history.Count - 1] == sheetName)
                return;

            _history.Add(sheetName);
            _index = _history.Count - 1;
        }

        public string GoBack()
        {
            if (!CanGoBack) return null;
            _index--;
            return _history[_index];
        }

        public string GoForward()
        {
            if (!CanGoForward) return null;
            _index++;
            return _history[_index];
        }

        public void Clear()
        {
            _history.Clear();
            _index = -1;
        }
    }
}