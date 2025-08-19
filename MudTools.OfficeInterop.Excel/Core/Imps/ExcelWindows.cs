//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel 窗口集合实现类
/// </summary>
internal class ExcelWindows : IExcelWindows
{
    private readonly MsExcel.Windows _windows;
    private readonly IExcelApplication _application;
    private readonly Dictionary<int, IExcelWindow> _windowCache;
    private bool _disposedValue;

    /// <summary>
    /// 获取窗口数量
    /// </summary>
    public int Count => _windows.Count;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _windows.Parent;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="windows">COM Windows 对象</param>
    /// <param name="application">关联的应用程序对象</param>
    internal ExcelWindows(MsExcel.Windows windows, IExcelApplication application)
    {
        _windows = windows ?? throw new ArgumentNullException(nameof(windows));
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _windowCache = new Dictionary<int, IExcelWindow>();
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取窗口（从1开始）
    /// </summary>
    /// <param name="index">窗口索引</param>
    /// <returns>窗口对象</returns>
    public IExcelWindow this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                // 检查缓存
                if (_windowCache.TryGetValue(index, out IExcelWindow cachedWindow))
                {
                    return cachedWindow;
                }

                var window = _windows.Item[index];
                var excelWindow = CreateExcelWindow(window);

                // 缓存窗口对象
                _windowCache[index] = excelWindow;
                return excelWindow;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get window at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 根据窗口标题获取窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <returns>窗口对象</returns>
    public IExcelWindow this[string caption]
    {
        get
        {
            object obj = Type.Missing;
            if (!string.IsNullOrEmpty(caption))
                obj = caption.Trim();
            var window = _windows.Item[obj];
            return CreateExcelWindow(window);
        }
    }

    /// <summary>
    /// 创建新窗口
    /// </summary>
    /// <returns>新创建的窗口对象</returns>
    public IExcelWindow NewWindow()
    {
        try
        {
            // 获取当前活动窗口
            var activeWindow = _windows.Parent as MsExcel.Application;
            if (activeWindow?.ActiveWindow == null)
                throw new InvalidOperationException("No active window found.");

            var newWindow = activeWindow.ActiveWindow.NewWindow();
            var excelWindow = CreateExcelWindow(newWindow);

            // 清除缓存以确保一致性
            _windowCache.Clear();

            return excelWindow;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create new window.", ex);
        }
    }

    /// <summary>
    /// 获取活动窗口
    /// </summary>
    /// <returns>活动窗口对象</returns>
    public IExcelWindow GetActiveWindow()
    {
        try
        {
            var activeWindow = (_windows.Parent as MsExcel.Application)?.ActiveWindow;
            if (activeWindow == null)
                return null;

            return CreateExcelWindow(activeWindow);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get active window.", ex);
        }
    }

    /// <summary>
    /// 根据条件查找窗口
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的窗口列表</returns>
    public IEnumerable<IExcelWindow> Find(Func<IExcelWindow, bool> predicate)
    {
        if (predicate == null)
            throw new ArgumentNullException(nameof(predicate));

        try
        {
            var results = new List<IExcelWindow>();
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
                    // 忽略获取失败的窗口
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

    /// <summary>
    /// 按窗口标题排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的窗口列表</returns>
    public IEnumerable<IExcelWindow> OrderByCaption(bool ascending = true)
    {
        try
        {
            var windows = new List<IExcelWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    windows.Add(this[i]);
                }
                catch
                {
                    // 忽略获取失败的窗口
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

    /// <summary>
    /// 按窗口索引排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的窗口列表</returns>
    public IEnumerable<IExcelWindow> OrderByIndex(bool ascending = true)
    {
        try
        {
            var windows = new List<IExcelWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    windows.Add(this[i]);
                }
                catch
                {
                    // 忽略获取失败的窗口
                    continue;
                }
            }

            return ascending
                ? windows.OrderBy(w => w.Index)
                : windows.OrderByDescending(w => w.Index);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to order windows by index.", ex);
        }
    }

    /// <summary>
    /// 刷新所有窗口
    /// </summary>
    public void RefreshAll()
    {
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    window.Refresh();
                }
                catch
                {
                    // 忽略刷新失败的窗口
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh all windows.", ex);
        }
    }

    /// <summary>
    /// 关闭所有窗口（除了指定的窗口）
    /// </summary>
    /// <param name="exceptWindow">要保留的窗口</param>
    public void CloseAllExcept(IExcelWindow? exceptWindow = null)
    {
        try
        {
            var windowsToClose = new List<IExcelWindow>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var window = this[i];
                    if (exceptWindow == null || window.Index != exceptWindow.Index)
                    {
                        windowsToClose.Add(window);
                    }
                }
                catch
                {
                    // 忽略获取失败的窗口
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
                    // 忽略关闭失败的窗口
                    continue;
                }
            }

            // 清除缓存
            _windowCache.Clear();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close windows.", ex);
        }
    }

    /// <summary>
    /// 激活所有窗口
    /// </summary>
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
                    // 忽略激活失败的窗口
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate all windows.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>窗口枚举器</returns>
    public IEnumerator<IExcelWindow> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 创建 ExcelWindow 对象
    /// </summary>
    /// <param name="window">COM Window 对象</param>
    /// <returns>ExcelWindow 对象</returns>
    private IExcelWindow CreateExcelWindow(MsExcel.Window window)
    {
        // 尝试找到关联的工作簿
        IExcelWorkbook associatedWorkbook = _application.ActiveWorkbook;
        return new ExcelWindow(window, associatedWorkbook);
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 清理缓存中的窗口对象
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

            // Windows 集合通常不需要显式释放 COM 对象
            // 因为它们是 Application 的子对象
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}