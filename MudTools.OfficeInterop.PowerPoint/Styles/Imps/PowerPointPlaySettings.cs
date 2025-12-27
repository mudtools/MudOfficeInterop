//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 播放设置实现类
/// </summary>
internal class PowerPointPlaySettings : IPowerPointPlaySettings
{
    private readonly MsPowerPoint.PlaySettings _playSettings;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置是否进入时播放
    /// </summary>
    public bool PlayOnEntry
    {
        get => _playSettings?.PlayOnEntry == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_playSettings != null)
                _playSettings.PlayOnEntry = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否隐藏未播放时
    /// </summary>
    public bool HideWhileNotPlaying
    {
        get => _playSettings?.HideWhileNotPlaying == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_playSettings != null)
                _playSettings.HideWhileNotPlaying = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否循环直到停止
    /// </summary>
    public bool LoopUntilStopped
    {
        get => _playSettings?.LoopUntilStopped == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_playSettings != null)
                _playSettings.LoopUntilStopped = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _playSettings?.Parent;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="playSettings">COM PlaySettings 对象</param>
    internal PowerPointPlaySettings(MsPowerPoint.PlaySettings playSettings)
    {
        _playSettings = playSettings;
        _disposedValue = false;
    }

    /// <summary>
    /// 设置播放参数
    /// </summary>
    /// <param name="playOnEntry">是否进入时播放</param>
    /// <param name="loopUntilStopped">是否循环直到停止</param>
    /// <param name="hideWhileNotPlaying">是否隐藏未播放时</param>
    public void SetPlaySettings(bool playOnEntry = true, bool loopUntilStopped = false, bool hideWhileNotPlaying = false)
    {
        try
        {
            if (_playSettings != null)
            {
                _playSettings.PlayOnEntry = playOnEntry ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                _playSettings.LoopUntilStopped = loopUntilStopped ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                _playSettings.HideWhileNotPlaying = hideWhileNotPlaying ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set play settings.", ex);
        }
    }


    /// <summary>
    /// 重置播放设置
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_playSettings != null)
            {
                _playSettings.PlayOnEntry = MsCore.MsoTriState.msoFalse;
                _playSettings.LoopUntilStopped = MsCore.MsoTriState.msoFalse;
                _playSettings.HideWhileNotPlaying = MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset play settings.", ex);
        }
    }

    /// <summary>
    /// 获取播放设置信息
    /// </summary>
    /// <returns>播放设置信息字符串</returns>
    public string GetPlaySettingsInfo()
    {
        try
        {
            return $"PlaySettings - PlayOnEntry: {PlayOnEntry}, LoopUntilStopped: {LoopUntilStopped}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get play settings info.", ex);
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