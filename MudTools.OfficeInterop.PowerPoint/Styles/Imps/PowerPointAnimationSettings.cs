//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 动画设置实现类
/// </summary>
internal class PowerPointAnimationSettings : IPowerPointAnimationSettings
{
    private readonly MsPowerPoint.AnimationSettings _animationSettings;
    private bool _disposedValue;
    private IPowerPointSoundEffect _soundEffect;
    private IPowerPointPlaySettings _playSettings;

    /// <summary>
    /// 获取或设置进入效果
    /// </summary>
    public int EntryEffect
    {
        get => _animationSettings != null ? (int)_animationSettings.EntryEffect : 0;
        set
        {
            if (_animationSettings != null)
                _animationSettings.EntryEffect = (MsPowerPoint.PpEntryEffect)value;
        }
    }

    /// <summary>
    /// 获取或设置动画顺序
    /// </summary>
    public int AnimationOrder
    {
        get => _animationSettings?.AnimationOrder ?? 0;
        set
        {
            if (_animationSettings != null)
                _animationSettings.AnimationOrder = value;
        }
    }

    /// <summary>
    /// 获取或设置前进模式
    /// </summary>
    public int AdvanceMode
    {
        get => _animationSettings != null ? (int)_animationSettings.AdvanceMode : 0;
        set
        {
            if (_animationSettings != null)
                _animationSettings.AdvanceMode = (MsPowerPoint.PpAdvanceMode)value;
        }
    }

    /// <summary>
    /// 获取或设置前进时间
    /// </summary>
    public float AdvanceTime
    {
        get => _animationSettings?.AdvanceTime ?? 0;
        set
        {
            if (_animationSettings != null)
                _animationSettings.AdvanceTime = value;
        }
    }

    /// <summary>
    /// 获取声音效果
    /// </summary>
    public IPowerPointSoundEffect SoundEffect
    {
        get
        {
            if (_soundEffect == null && _animationSettings?.SoundEffect != null)
            {
                _soundEffect = new PowerPointSoundEffect(_animationSettings.SoundEffect);
            }
            return _soundEffect;
        }
    }

    /// <summary>
    /// 获取播放设置
    /// </summary>
    public IPowerPointPlaySettings PlaySettings
    {
        get
        {
            if (_playSettings == null && _animationSettings?.PlaySettings != null)
            {
                _playSettings = new PowerPointPlaySettings(_animationSettings.PlaySettings);
            }
            return _playSettings;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _animationSettings?.Parent;


    /// <summary>
    /// 获取或设置是否动画背景
    /// </summary>
    public bool AnimateBackground
    {
        get => _animationSettings?.AnimateBackground == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_animationSettings != null)
                _animationSettings.AnimateBackground = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }




    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="animationSettings">COM AnimationSettings 对象</param>
    internal PowerPointAnimationSettings(MsPowerPoint.AnimationSettings animationSettings)
    {
        _animationSettings = animationSettings;
        _disposedValue = false;
    }

    /// <summary>
    /// 播放动画
    /// </summary>
    /// <param name="from">起始时间</param>
    /// <param name="to">结束时间</param>
    /// <param name="repeats">重复次数</param>
    public void Play(double from = 0, double to = 0, int repeats = 1)
    {
        try
        {
            // 动画播放需要通过幻灯片放映实现
            throw new NotImplementedException("Playing animation is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to play animation.", ex);
        }
    }

    /// <summary>
    /// 停止动画
    /// </summary>
    public void Stop()
    {
        try
        {
            // 动画停止需要通过幻灯片放映实现
            throw new NotImplementedException("Stopping animation is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to stop animation.", ex);
        }
    }

    /// <summary>
    /// 暂停动画
    /// </summary>
    public void Pause()
    {
        try
        {
            // 动画暂停需要通过幻灯片放映实现
            throw new NotImplementedException("Pausing animation is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to pause animation.", ex);
        }
    }

    /// <summary>
    /// 恢复动画
    /// </summary>
    public void Resume()
    {
        try
        {
            // 动画恢复需要通过幻灯片放映实现
            throw new NotImplementedException("Resuming animation is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to resume animation.", ex);
        }
    }

    /// <summary>
    /// 重置动画设置
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_animationSettings != null)
            {
                _animationSettings.EntryEffect = MsPowerPoint.PpEntryEffect.ppEffectNone;
                _animationSettings.AdvanceMode = MsPowerPoint.PpAdvanceMode.ppAdvanceOnClick;
                _animationSettings.AdvanceTime = 0;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset animation settings.", ex);
        }
    }

    /// <summary>
    /// 应用动画方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    public void ApplyAnimationScheme(int schemeIndex = -1)
    {
        try
        {
            // 动画方案应用需要通过幻灯片对象实现
            throw new NotImplementedException("Applying animation scheme is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply animation scheme.", ex);
        }
    }

    /// <summary>
    /// 设置动画参数
    /// </summary>
    /// <param name="entryEffect">进入效果</param>
    /// <param name="advanceMode">前进模式</param>
    /// <param name="advanceTime">前进时间</param>
    public void SetAnimation(int entryEffect = 0, int advanceMode = 1, float advanceTime = 0)
    {
        try
        {
            if (_animationSettings != null)
            {
                _animationSettings.EntryEffect = (MsPowerPoint.PpEntryEffect)entryEffect;
                _animationSettings.AdvanceMode = (MsPowerPoint.PpAdvanceMode)advanceMode;
                _animationSettings.AdvanceTime = advanceTime;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set animation parameters.", ex);
        }
    }

    /// <summary>
    /// 获取动画设置信息
    /// </summary>
    /// <returns>动画设置信息字符串</returns>
    public string GetAnimationSettingsInfo()
    {
        try
        {
            return $"AnimationSettings - EntryEffect: {EntryEffect}, AdvanceMode: {AdvanceMode}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get animation settings info.", ex);
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
            _soundEffect?.Dispose();
            _playSettings?.Dispose();
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
