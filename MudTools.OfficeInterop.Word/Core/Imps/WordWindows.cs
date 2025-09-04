//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 窗口集合实现类
/// </summary>
internal class WordWindows : IWordWindows
{
    private readonly MsWord.Windows _windows;
    private readonly Dictionary<int, IWordWindow> _windowCache;
    private bool _disposedValue;

    public int Count => _windows.Count;

    public object Parent => _windows.Parent;


    internal WordWindows(MsWord.Windows windows)
    {
        _windows = windows ?? throw new ArgumentNullException(nameof(windows));
        _windowCache = new Dictionary<int, IWordWindow>();
        _disposedValue = false;
    }

    public IWordWindow this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                if (_windowCache.TryGetValue(index, out IWordWindow cachedWindow))
                {
                    return cachedWindow;
                }

                var window = _windows[index];
                var wordWindow = new WordWindow(window);
                _windowCache[index] = wordWindow;
                return wordWindow;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get window at index {index}.", ex);
            }
        }
    }

    public IWordWindow this[string caption]
    {
        get
        {
            if (string.IsNullOrEmpty(caption))
                throw new ArgumentException("Caption cannot be null or empty.", nameof(caption));

            try
            {
                var window = _windows[caption];
                return new WordWindow(window);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get window with caption '{caption}'.", ex);
            }
        }
    }

    public IWordWindow NewWindow()
    {
        try
        {
            var activeWindow = _windows.Parent as MsWord.Application;
            if (activeWindow?.ActiveWindow == null)
                throw new InvalidOperationException("No active window found.");

            var newWindow = activeWindow.ActiveWindow.NewWindow();
            var wordWindow = new WordWindow(newWindow);
            _windowCache.Clear();
            return wordWindow;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create new window.", ex);
        }
    }

    public IWordWindow GetActiveWindow()
    {
        try
        {
            var activeWindow = (_windows.Parent as MsWord.Application)?.ActiveWindow;
            if (activeWindow == null)
                return null;

            return new WordWindow(activeWindow);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get active window.", ex);
        }
    }

    public IEnumerable<IWordWindow> Find(Func<IWordWindow, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        try
        {
            var results = new List<IWordWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    if (predicate(window))
                    {
                        results.Add(window);
                    }
                }
                catch
                {
                    continue;
                }
            }
            return results;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find windows.", ex);
        }
    }

    public IEnumerable<IWordWindow> OrderByCaption(bool ascending = true)
    {
        try
        {
            var windows = new List<IWordWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    windows.Add(this[i]);
                }
                catch
                {
                    continue;
                }
            }

            return ascending
                ? windows.OrderBy(w => w.Caption, StringComparer.OrdinalIgnoreCase)
                : windows.OrderByDescending(w => w.Caption, StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to order windows by caption.", ex);
        }
    }

    public IEnumerable<IWordWindow> OrderByIndex(bool ascending = true)
    {
        try
        {
            var windows = new List<IWordWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    windows.Add(this[i]);
                }
                catch
                {
                    continue;
                }
            }

            return ascending
                ? windows.OrderBy(w => w.GetHashCode())
                : windows.OrderByDescending(w => w.GetHashCode());
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to order windows by index.", ex);
        }
    }

    public void RefreshAll()
    {
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    // Window 对象没有直接的刷新方法
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh all windows.", ex);
        }
    }

    public void CloseAllExcept(IWordWindow exceptWindow = null)
    {
        try
        {
            var windowsToClose = new List<IWordWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    if (exceptWindow == null || window.GetHashCode() != exceptWindow.GetHashCode())
                    {
                        windowsToClose.Add(window);
                    }
                }
                catch
                {
                    continue;
                }
            }

            foreach (var window in windowsToClose)
            {
                try
                {
                    window.Close();
                }
                catch
                {
                    continue;
                }
            }

            _windowCache.Clear();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close windows.", ex);
        }
    }

    public void ActivateAll()
    {
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    window.Activate();
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate all windows.", ex);
        }
    }

    public IEnumerator<IWordWindow> GetEnumerator()
    {
        try
        {
            var windows = new List<IWordWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    windows.Add(this[i]);
                }
                catch
                {
                    continue;
                }
            }
            return windows.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate windows.", ex);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            foreach (var window in _windowCache.Values)
            {
                try
                {
                    window?.Dispose();
                }
                catch
                {
                    // 忽略释放失败的情况
                }
            }
            _windowCache.Clear();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}