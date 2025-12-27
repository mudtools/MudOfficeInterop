//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 幻灯片放映设置实现类
/// </summary>
internal class PowerPointSlideShowSettings : IPowerPointSlideShowSettings
{
    private readonly MsPowerPoint.SlideShowSettings _slideShowSettings;
    private bool _disposedValue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _slideShowSettings.Parent;

    /// <summary>
    /// 获取或设置起始幻灯片索引
    /// </summary>
    public int StartingSlide
    {
        get => _slideShowSettings.StartingSlide;
        set => _slideShowSettings.StartingSlide = value;
    }

    /// <summary>
    /// 获取或设置结束幻灯片索引
    /// </summary>
    public int EndingSlide
    {
        get => _slideShowSettings.EndingSlide;
        set => _slideShowSettings.EndingSlide = value;
    }

    /// <summary>
    /// 获取或设置是否循环放映
    /// </summary>
    public bool LoopUntilStopped
    {
        get => _slideShowSettings.LoopUntilStopped == Microsoft.Office.Core.MsoTriState.msoTrue;
        set => _slideShowSettings.LoopUntilStopped = value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }


    /// <summary>
    /// 获取或设置是否显示滚动条
    /// </summary>
    public bool ShowScrollbar
    {
        get => _slideShowSettings.ShowScrollbar == Microsoft.Office.Core.MsoTriState.msoTrue;
        set => _slideShowSettings.ShowScrollbar = value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="slideShowSettings">COM SlideShowSettings 对象</param>
    internal PowerPointSlideShowSettings(MsPowerPoint.SlideShowSettings slideShowSettings)
    {
        _slideShowSettings = slideShowSettings ?? throw new ArgumentNullException(nameof(slideShowSettings));
        _disposedValue = false;
    }

    /// <summary>
    /// 运行幻灯片放映
    /// </summary>
    /// <returns>幻灯片放映窗口</returns>
    public IPowerPointSlideShowWindow Run()
    {
        try
        {
            var slideShowWindow = _slideShowSettings.Run();
            return new PowerPointSlideShowWindow(slideShowWindow);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to run slide show.", ex);
        }
    }

    /// <summary>
    /// 应用设置
    /// </summary>
    public void Apply()
    {
        try
        {
            // SlideShowSettings 的设置通常是自动应用的
            // 这里作为占位符实现
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply slide show settings.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
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
