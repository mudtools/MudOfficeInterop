//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 视图实现类
/// </summary>
internal class PowerPointView : IPowerPointView
{
    private MsPowerPoint.View _view;
    private bool _disposedValue;
    private IPowerPointSlideShowWindow _slideShowWindow;
    private IPowerPointPresentation _presentation;

    /// <summary>
    /// 获取视图类型
    /// </summary>
    public PpViewType Type => (PpViewType)_view.Type;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _view.Parent;


    /// <summary>
    /// 获取当前幻灯片
    /// </summary>
    public IPowerPointSlide Slide
    {
        get
        {
            try
            {
                var slide = _view.Slide as MsPowerPoint.Slide;
                return slide != null ? new PowerPointSlide(slide) : null;
            }
            catch
            {
                return null;
            }
        }
    }


    /// <summary>
    /// 获取或设置缩放比例
    /// </summary>
    public int Zoom
    {
        get => _view.Zoom;
        set => _view.Zoom = value;
    }


    /// <summary>
    /// 获取选中的形状范围
    /// </summary>
    public IPowerPointShapeRange Selection
    {
        get
        {
            try
            {
                var selection = _view.Application.ActiveWindow.Selection;
                if (selection?.ShapeRange != null)
                {
                    // 这里需要具体的 ShapeRange 实现
                    return null; // 占位符
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="view">COM View 对象</param>
    /// <param name="presentation">关联的演示文稿对象</param>
    internal PowerPointView(MsPowerPoint.View view, IPowerPointPresentation presentation)
    {
        _view = view ?? throw new ArgumentNullException(nameof(view));
        _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
        _disposedValue = false;
    }

    /// <summary>
    /// 转到指定幻灯片
    /// </summary>
    /// <param name="slideIndex">幻灯片索引</param>
    public void GoToSlide(int slideIndex)
    {
        if (slideIndex < 1)
            throw new ArgumentOutOfRangeException(nameof(slideIndex), "Slide index must be greater than 0.");

        try
        {
            _view.GotoSlide(slideIndex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to go to slide {slideIndex}.", ex);
        }
    }


    /// <summary>
    /// 开始幻灯片放映
    /// </summary>
    /// <param name="fromSlide">起始幻灯片索引</param>
    /// <param name="toSlide">结束幻灯片索引</param>
    /// <returns>幻灯片放映窗口对象</returns>
    public IPowerPointSlideShowWindow StartSlideShow(int fromSlide = 1, int toSlide = -1)
    {
        try
        {
            var slideShowSettings = _view.Application.ActivePresentation.SlideShowSettings;
            slideShowSettings.StartingSlide = fromSlide;
            slideShowSettings.EndingSlide = toSlide > 0 ? toSlide : _view.Application.ActivePresentation.Slides.Count;

            var slideShowWindow = slideShowSettings.Run();
            var window = new PowerPointSlideShowWindow(slideShowWindow);
            _slideShowWindow = window;
            return window;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to start slide show.", ex);
        }
    }

    /// <summary>
    /// 激活当前视图
    /// </summary>
    public void Activate()
    {
        try
        {
            _view.Application.ActiveWindow.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate view.", ex);
        }
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
            try
            {
                // 释放底层COM对象
                if (_view != null)
                    Marshal.ReleaseComObject(_view);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _view = null;
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