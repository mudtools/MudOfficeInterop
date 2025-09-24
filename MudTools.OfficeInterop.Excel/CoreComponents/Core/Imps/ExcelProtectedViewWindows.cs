//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Vbe.Interop;

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelProtectedViewWindows : IExcelProtectedViewWindows
{
    private DisposableList _disposables = [];
    private MsExcel.ProtectedViewWindows? _protectedViewWindows;
    private bool _disposedValue;

    public int Count => _protectedViewWindows != null ? _protectedViewWindows.Count : 0;

    public IExcelProtectedViewWindow? this[int index]
    {
        get
        {
            if (_protectedViewWindows == null)
                return null;

            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index));

            var win = new ExcelProtectedViewWindow(_protectedViewWindows[index]);
            _disposables.Add(win);
            return win;
        }
    }

    public IExcelProtectedViewWindow? this[string caption]
    {
        get
        {
            if (_protectedViewWindows == null)
                return null;

            var win = new ExcelProtectedViewWindow(_protectedViewWindows[caption]);
            _disposables.Add(win);
            return win;
        }
    }

    internal ExcelProtectedViewWindows(MsExcel.ProtectedViewWindows protectedViewWindows)
    {
        _protectedViewWindows = protectedViewWindows ?? throw new ArgumentNullException(nameof(protectedViewWindows));
        _disposedValue = false;
    }

    public IExcelProtectedViewWindow? Open(string filename, string password = null,
                                         bool readOnlyRecommended = false, bool editable = false)
    {
        if (string.IsNullOrEmpty(filename))
            throw new ArgumentException("文件路径不能为空。", nameof(filename));
        if (_protectedViewWindows == null)
            return null;
        try
        {
            var window = _protectedViewWindows.Open(filename, password ?? string.Empty,
                                                   readOnlyRecommended, editable);
            return window != null ? new ExcelProtectedViewWindow(window) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {filename}", ex);
        }
        catch (Exception x)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {filename}", x);
        }
    }

    /// <inheritdoc/>
    public IExcelProtectedViewWindow? FindByFilename(string filename)
    {
        if (string.IsNullOrEmpty(filename))
            throw new ArgumentException("文件路径不能为空。", nameof(filename));
        if (_protectedViewWindows == null)
            return null;
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var window = this[i];
                if (string.Equals(window.SourceName, filename, StringComparison.OrdinalIgnoreCase))
                {
                    return window;
                }
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {filename}", ex);
        }
        catch (Exception x)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {filename}", x);
        }

    }

    public IExcelProtectedViewWindow FindByCaption(string caption)
    {
        if (string.IsNullOrEmpty(caption))
            throw new ArgumentException("窗口标题不能为空。", nameof(caption));

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var window = this[i];
                if (string.Equals(window.Caption, caption, StringComparison.OrdinalIgnoreCase))
                {
                    return window;
                }
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {caption}", ex);
        }
        catch (Exception x)
        {
            throw new InvalidOperationException($"无法打开文件到受保护视图: {caption}", x);
        }
    }

    public IExcelApplication? Parent => _protectedViewWindows != null ? new ExcelApplication(_protectedViewWindows.Application) : null;

    public IExcelProtectedViewWindow? ActiveProtectedViewWindow => _protectedViewWindows != null ? new ExcelProtectedViewWindow(_protectedViewWindows.Application.ActiveProtectedViewWindow) : null;

    public IEnumerable<IExcelProtectedViewWindow> VisibleWindows
    {
        get
        {
            var result = new List<IExcelProtectedViewWindow>();
            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    var window = this[i];
                    if (window.Visible)
                    {
                        result.Add(window);
                    }
                }
            }
            catch (COMException ce)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", ce);
            }
            catch (Exception x)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", x);
            }
            return result;
        }
    }

    public IEnumerable<IExcelProtectedViewWindow> MaximizedWindows
    {
        get
        {
            var result = new List<IExcelProtectedViewWindow>();
            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    var window = this[i];
                    if (window.WindowState == XlProtectedViewWindowState.xlProtectedViewWindowMaximized)
                    {
                        result.Add(window);
                    }
                }
            }
            catch (COMException ce)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", ce);
            }
            catch (Exception x)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", x);
            }
            return result;
        }
    }

    public IEnumerable<IExcelProtectedViewWindow> MinimizedWindows
    {
        get
        {
            var result = new List<IExcelProtectedViewWindow>();
            try
            {
                for (int i = 1; i <= Count; i++)
                {
                    var window = this[i];
                    if (window.WindowState == XlProtectedViewWindowState.xlProtectedViewWindowMinimized)
                    {
                        result.Add(window);
                    }
                }
            }
            catch (COMException ce)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", ce);
            }
            catch (Exception x)
            {
                throw new InvalidOperationException("无法获取受保护视图窗口。", x);
            }
            return result;
        }
    }

    public IEnumerable<IExcelProtectedViewWindow> GetWindowsByState(XlProtectedViewWindowState state)
    {
        var result = new List<IExcelProtectedViewWindow>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var window = this[i];
                if (window.WindowState == state)
                {
                    result.Add(window);
                }
            }
        }
        catch (COMException ce)
        {
            throw new InvalidOperationException("无法获取受保护视图窗口。", ce);
        }
        catch (Exception x)
        {
            throw new InvalidOperationException("无法获取受保护视图窗口。", x);
        }
        return result;
    }

    public IEnumerator<IExcelProtectedViewWindow> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _protectedViewWindows != null)
        {
            _disposables?.Dispose();
            Marshal.ReleaseComObject(_protectedViewWindows);
            _protectedViewWindows = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}