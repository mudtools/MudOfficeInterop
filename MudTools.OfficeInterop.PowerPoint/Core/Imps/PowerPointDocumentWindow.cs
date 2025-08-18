//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint DocumentWindow 对象的二次封装实现类
/// 实现 IPowerPointWindow 接口
/// </summary>
internal class PowerPointDocumentWindow : IPowerPointDocumentWindow
{
    private MsPowerPoint.DocumentWindow _window;
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 PowerPointWindow 实例
    /// </summary>
    /// <param name="window">要封装的 Microsoft.Office.Interop.PowerPoint.DocumentWindow 对象</param>
    internal PowerPointDocumentWindow(MsPowerPoint.DocumentWindow window)
    {
        _window = window ?? throw new ArgumentNullException(nameof(window));
    }

    #region 基础属性
    public int? Hwnd => _window.HWND;
    public object Parent => _window.Parent;

    public IPowerPointApplication Application => _window.Application != null ? new PowerPointApplication(_window.Application) : null;

    public IPowerPointPresentation Presentation => _window.Presentation != null ? new PowerPointPresentation(_window.Presentation) : null;

    public string Caption
    {
        get => _window.Caption;
    }

    public bool IsActive => _window.Active == MsCore.MsoTriState.msoTrue;
    #endregion

    #region 位置和大小
    public float Left
    {
        get => _window.Left;
        set => _window.Left = value;
    }

    public float Top
    {
        get => _window.Top;
        set => _window.Top = value;
    }

    public float Width
    {
        get => _window.Width;
        set => _window.Width = value;
    }

    public float Height
    {
        get => _window.Height;
        set => _window.Height = value;
    }

    public PpWindowState WindowState
    {
        get => (PpWindowState)_window.WindowState;
        set => _window.WindowState = (MsPowerPoint.PpWindowState)value;
    }
    #endregion

    #region 核心对象
    public IPowerPointSelection Selection => _window.Selection != null ? new PowerPointSelection(_window.Selection) : null;

    public IPowerPointView View => _window.View != null ? new PowerPointView(_window.View, this.Presentation) : null;

    public IPowerPointSlide ActiveSlide
    {
        get
        {
            try
            {
                if (_window.View?.Type == MsPowerPoint.PpViewType.ppViewNormal && _window.View is MsPowerPoint.SlideShowView slideView)
                {
                    return slideView.Slide != null ? new PowerPointSlide(slideView.Slide) : null;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
    #endregion

    #region 操作方法
    public void Activate()
    {
        _window.Activate();
    }

    public void Close()
    {
        try
        {
            _window.Close();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error closing window: {ex.Message}");
        }
    }
    #endregion

    #region 视图操作
    public void ViewNormal()
    {
        try { _window.ViewType = MsPowerPoint.PpViewType.ppViewNormal; } catch { }
    }

    public void ViewSlideSorter()
    {
        try { _window.ViewType = MsPowerPoint.PpViewType.ppViewSlideSorter; } catch { }
    }

    public void ViewSlideShow()
    {
        try { _window.ViewType = MsPowerPoint.PpViewType.ppViewSlide; } catch { }
    }

    public void ViewNotesPage()
    {
        try { _window.ViewType = MsPowerPoint.PpViewType.ppViewNotesPage; } catch { }
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

    ~PowerPointDocumentWindow()
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
