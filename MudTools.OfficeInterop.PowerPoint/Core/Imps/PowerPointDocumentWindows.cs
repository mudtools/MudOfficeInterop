//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint DocumentWindows 集合对象的二次封装实现类
/// 实现 IPowerPointWindows 接口
/// </summary>
internal class PowerPointDocumentWindows : IPowerPointDocumentWindows
{
    private MsPowerPoint.DocumentWindows _windows;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 PowerPointWindows 实例
    /// </summary>
    /// <param name="windows">要封装的 Microsoft.Office.Interop.PowerPoint.DocumentWindows 对象</param>
    internal PowerPointDocumentWindows(MsPowerPoint.DocumentWindows windows)
    {
        _windows = windows ?? throw new ArgumentNullException(nameof(windows));
    }

    #region 基础属性
    public int Count => _windows.Count;

    public IPowerPointDocumentWindow this[int index] => new PowerPointDocumentWindow(_windows[index]);

    public object? Parent => _windows.Parent;

    public IPowerPointApplication Application => _windows.Application != null ? new PowerPointApplication(_windows.Application) : null;
    #endregion

    #region 查找和筛选
    public IPowerPointDocumentWindow[] FindByCaption(string caption, bool matchCase = false)
    {
        var results = new List<IPowerPointDocumentWindow>();
        for (int i = 1; i <= Count; i++)
        {
            var window = this[i];
            if (string.Compare(window.Caption, caption, !matchCase) == 0)
                results.Add(window);
        }
        return results.ToArray();
    }

    public IPowerPointDocumentWindow[] FindByPresentationName(string presentationName, bool matchCase = false)
    {
        var results = new List<IPowerPointDocumentWindow>();
        for (int i = 1; i <= Count; i++)
        {
            var window = this[i];
            if (string.Compare(window.Presentation?.Name, presentationName, !matchCase) == 0) // Requires Presentation.Name
            {
                results.Add(window);
            }
        }
        return results.ToArray();
    }

    public IPowerPointDocumentWindow GetActiveWindow()
    {
        try
        {
            var activeWin = _windows.Application.ActiveWindow;
            if (activeWin != null)
            {
                return new PowerPointDocumentWindow(activeWin);
            }
            return null;
        }
        catch
        {
            // Handle error
        }
        return null;
    }

    public IPowerPointDocumentWindow[] GetVisibleWindows()
    {
        var results = new List<IPowerPointDocumentWindow>();
        for (int i = 1; i <= Count; i++)
        {
            var window = this[i];
            if (window.WindowState != PpWindowState.ppWindowMinimized)
            {
                results.Add(window);
            }
        }
        return results.ToArray();
    }

    #endregion

    #region 操作方法
    public void Clear()
    {
        // Close all windows. Iterate backwards to avoid index issues.
        for (int i = Count; i >= 1; i--)
        {
            try
            {
                Delete(i);
            }
            catch
            {

            }
        }
    }

    public void Delete(int index)
    {
        try
        {
            MsPowerPoint.DocumentWindow window = _windows[index];
            window?.Close();
        }
        catch
        {

        }
    }

    public void Delete(IPowerPointDocumentWindow window)
    {
        if (window is PowerPointDocumentWindow pptWindow)
        {
            try
            {
                pptWindow.Close();
            }
            catch { }
        }
    }

    public void DeleteRange(int[] indices)
    {
        // Sort indices descending to avoid index shifting issues
        var sortedIndices = new List<int>(indices);
        sortedIndices.Sort((a, b) => b.CompareTo(a)); // Descending sort
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }
    #endregion

    #region IEnumerable<IPowerPointWindow> Support
    public IEnumerator<IPowerPointDocumentWindow> GetEnumerator()
    {
        for (int i = 1; i <= _windows.Count; i++)
        {
            yield return new PowerPointDocumentWindow(_windows[i]);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _disposedValue = true;
        }
    }

    ~PowerPointDocumentWindows()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
