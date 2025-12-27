//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;


/// <summary>
/// PowerPoint 幻灯片放映视图实现类
/// </summary>
internal class PowerPointSlideShowView : IPowerPointSlideShowView
{
    private readonly MsPowerPoint.SlideShowView _slideShowView;
    private bool _disposedValue;
    private IPowerPointSlide _slide;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _slideShowView.Parent;

    /// <summary>
    /// 获取当前幻灯片
    /// </summary>
    public IPowerPointSlide Slide
    {
        get
        {
            if (_slide == null)
            {
                _slide = new PowerPointSlide(_slideShowView.Slide);
            }
            return _slide;
        }
    }

    /// <summary>
    /// 获取当前幻灯片索引
    /// </summary>
    public int SlideIndex => _slideShowView.CurrentShowPosition;

    /// <summary>
    /// 获取幻灯片放映状态
    /// </summary>
    public PpSlideShowState State => (PpSlideShowState)_slideShowView.State;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="slideShowView">COM SlideShowView 对象</param>
    internal PowerPointSlideShowView(MsPowerPoint.SlideShowView slideShowView)
    {
        _slideShowView = slideShowView ?? throw new ArgumentNullException(nameof(slideShowView));
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
            _slideShowView.GotoSlide(slideIndex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to go to slide {slideIndex}.", ex);
        }
    }

    /// <summary>
    /// 转到下一张幻灯片
    /// </summary>
    public void NextSlide()
    {
        try
        {
            _slideShowView.Next();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to go to next slide.", ex);
        }
    }

    /// <summary>
    /// 转到上一张幻灯片
    /// </summary>
    public void PreviousSlide()
    {
        try
        {
            _slideShowView.Previous();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to go to previous slide.", ex);
        }
    }

    /// <summary>
    /// 转到第一张幻灯片
    /// </summary>
    public void FirstSlide()
    {
        try
        {
            _slideShowView.First();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to go to first slide.", ex);
        }
    }

    /// <summary>
    /// 转到最后一张幻灯片
    /// </summary>
    public void LastSlide()
    {
        try
        {
            _slideShowView.Last();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to go to last slide.", ex);
        }
    }


    /// <summary>
    /// 结束幻灯片放映
    /// </summary>
    public void End()
    {
        try
        {
            _slideShowView.Exit();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to end slide show.", ex);
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
            _slide?.Dispose();
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