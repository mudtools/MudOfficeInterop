//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelProtectedViewWindow : IExcelProtectedViewWindow
{
    private MsExcel.ProtectedViewWindow _protectedViewWindow;
    private bool _disposedValue;

    public string Caption
    {
        get => _protectedViewWindow.Caption;
        set => _protectedViewWindow.Caption = value;
    }

    public bool Visible
    {
        get => _protectedViewWindow.Visible;
        set => _protectedViewWindow.Visible = value;
    }

    public bool EnableResize
    {
        get => _protectedViewWindow.EnableResize;
        set => _protectedViewWindow.EnableResize = value;
    }

    public double Left
    {
        get => _protectedViewWindow.Left;
        set => _protectedViewWindow.Left = value;
    }

    public double Top
    {
        get => _protectedViewWindow.Top;
        set => _protectedViewWindow.Top = value;
    }

    public double Width
    {
        get => _protectedViewWindow.Width;
        set => _protectedViewWindow.Width = value;
    }

    public double Height
    {
        get => _protectedViewWindow.Height;
        set => _protectedViewWindow.Height = value;
    }

    public XlProtectedViewWindowState WindowState
    {
        get => (XlProtectedViewWindowState)(int)_protectedViewWindow.WindowState;
        set => _protectedViewWindow.WindowState = (MsExcel.XlProtectedViewWindowState)value;
    }


    public IExcelWorkbook Workbook => new ExcelWorkbook(_protectedViewWindow.Workbook);

    public string SourcePath => _protectedViewWindow.SourcePath;

    public string SourceName => _protectedViewWindow.SourceName;


    internal ExcelProtectedViewWindow(MsExcel.ProtectedViewWindow protectedViewWindow)
    {
        _protectedViewWindow = protectedViewWindow ?? throw new ArgumentNullException(nameof(protectedViewWindow));
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _protectedViewWindow.Activate();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法激活受保护视图窗口。", ex);
        }
    }

    public void Close()
    {
        try
        {
            _protectedViewWindow.Close();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法关闭受保护视图窗口。", ex);
        }
    }

    public IExcelWorkbook Edit()
    {
        try
        {
            var workbook = _protectedViewWindow.Edit();
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法编辑受保护视图中的工作簿。", ex);
        }
    }

    public void Maximize()
    {
        try
        {
            _protectedViewWindow.WindowState = MsExcel.XlProtectedViewWindowState.xlProtectedViewWindowMaximized;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法最大化受保护视图窗口。", ex);
        }
    }

    public void Minimize()
    {
        try
        {
            _protectedViewWindow.WindowState = MsExcel.XlProtectedViewWindowState.xlProtectedViewWindowMinimized;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法最小化受保护视图窗口。", ex);
        }
    }

    public void Restore()
    {
        try
        {
            _protectedViewWindow.WindowState = MsExcel.XlProtectedViewWindowState.xlProtectedViewWindowNormal;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法恢复受保护视图窗口到正常大小。", ex);
        }
    }

    public void Move(int left, int top)
    {
        try
        {
            _protectedViewWindow.Left = left;
            _protectedViewWindow.Top = top;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法移动受保护视图窗口。", ex);
        }
    }

    public void Resize(int width, int height)
    {
        try
        {
            _protectedViewWindow.Width = width;
            _protectedViewWindow.Height = height;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法调整受保护视图窗口大小。", ex);
        }
    }


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _protectedViewWindow != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_protectedViewWindow) > 0) { }
            }
            catch { }
            _protectedViewWindow = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}