//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 幻灯片放映窗口实现类
/// </summary>
internal class PowerPointSlideShowWindow : IPowerPointSlideShowWindow
{
    private readonly MsPowerPoint.SlideShowWindow _slideShowWindow;
    private bool _disposedValue;
    private IPowerPointSlideShowView _view;
    private IPowerPointSlideShowSettings _settings;


    /// <summary>
    /// 获取或设置窗口高度
    /// </summary>
    public float Height
    {
        get => _slideShowWindow.Height;
        set => _slideShowWindow.Height = value;
    }

    /// <summary>
    /// 获取或设置窗口宽度
    /// </summary>
    public float Width
    {
        get => _slideShowWindow.Width;
        set => _slideShowWindow.Width = value;
    }

    /// <summary>
    /// 获取或设置窗口左边缘位置
    /// </summary>
    public float Left
    {
        get => _slideShowWindow.Left;
        set => _slideShowWindow.Left = value;
    }

    /// <summary>
    /// 获取或设置窗口上边缘位置
    /// </summary>
    public float Top
    {
        get => _slideShowWindow.Top;
        set => _slideShowWindow.Top = value;
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _slideShowWindow.Parent;

    /// <summary>
    /// 获取幻灯片放映视图
    /// </summary>
    public IPowerPointSlideShowView View
    {
        get
        {
            if (_view == null)
            {
                _view = new PowerPointSlideShowView(_slideShowWindow.View);
            }
            return _view;
        }
    }

    /// <summary>
    /// 获取幻灯片放映设置
    /// </summary>
    public IPowerPointSlideShowSettings Settings
    {
        get
        {
            if (_settings == null)
            {
                _settings = new PowerPointSlideShowSettings(_slideShowWindow.Presentation.SlideShowSettings);
            }
            return _settings;
        }
    }

    /// <summary>
    /// 获取幻灯片放映状态
    /// </summary>
    public PpSlideShowState State => (PpSlideShowState)_slideShowWindow.View.State;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="slideShowWindow">COM SlideShowWindow 对象</param>
    internal PowerPointSlideShowWindow(MsPowerPoint.SlideShowWindow slideShowWindow)
    {
        _slideShowWindow = slideShowWindow ?? throw new ArgumentNullException(nameof(slideShowWindow));
        _disposedValue = false;
    }

    /// <summary>
    /// 激活窗口
    /// </summary>
    public void Activate()
    {
        try
        {
            _slideShowWindow.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate slide show window.", ex);
        }
    }

    /// <summary>
    /// 暂停幻灯片放映
    /// </summary>
    public void Pause()
    {
        try
        {
            _slideShowWindow.View.State = MsPowerPoint.PpSlideShowState.ppSlideShowPaused;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to pause slide show.", ex);
        }
    }

    /// <summary>
    /// 恢复幻灯片放映
    /// </summary>
    public void Resume()
    {
        try
        {
            _slideShowWindow.View.State = MsPowerPoint.PpSlideShowState.ppSlideShowRunning;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to resume slide show.", ex);
        }
    }

    /// <summary>
    /// 切换到黑屏
    /// </summary>
    public void BlackScreen()
    {
        try
        {
            _slideShowWindow.View.State = MsPowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to switch to black screen.", ex);
        }
    }

    /// <summary>
    /// 切换到白屏
    /// </summary>
    public void WhiteScreen()
    {
        try
        {
            _slideShowWindow.View.State = MsPowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to switch to white screen.", ex);
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
            _view?.Dispose();
            _settings?.Dispose();
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