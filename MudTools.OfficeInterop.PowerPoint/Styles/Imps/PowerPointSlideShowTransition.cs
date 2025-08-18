//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 幻灯片放映切换效果实现类
/// </summary>
internal class PowerPointSlideShowTransition : IPowerPointSlideShowTransition
{
    private readonly MsPowerPoint.SlideShowTransition _transition;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置进入效果
    /// </summary>
    public int EntryEffect
    {
        get => _transition != null ? (int)_transition.EntryEffect : 0;
        set
        {
            if (_transition != null)
                _transition.EntryEffect = (MsPowerPoint.PpEntryEffect)value;
        }
    }

    /// <summary>
    /// 获取或设置是否定时前进
    /// </summary>
    public int AdvanceOnTime
    {
        get => _transition != null ? (int)_transition.AdvanceOnTime : 0;
        set
        {
            if (_transition != null)
                _transition.AdvanceOnTime = (MsoTriState)value;
        }
    }

    /// <summary>
    /// 获取或设置前进时间
    /// </summary>
    public float AdvanceTime
    {
        get => _transition?.AdvanceTime ?? 0;
        set
        {
            if (_transition != null)
                _transition.AdvanceTime = value;
        }
    }

    /// <summary>
    /// 获取或设置是否隐藏幻灯片
    /// </summary>
    public bool Hidden
    {
        get => _transition.Hidden == MsoTriState.msoCTrue;
        set
        {
            if (_transition != null)
                _transition.Hidden = value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _transition?.Parent;

    /// <summary>
    /// 获取或设置声音效果
    /// </summary>
    public string SoundEffect
    {
        get
        {
            try
            {
                return _transition?.SoundEffect?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        set
        {
            try
            {
                if (_transition?.SoundEffect != null)
                {
                    _transition.SoundEffect.ImportFromFile(value ?? string.Empty);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to set sound effect '{value}'.", ex);
            }
        }
    }

    /// <summary>
    /// 获取或设置持续时间
    /// </summary>
    public float Duration
    {
        get => _transition?.Duration ?? 0;
        set
        {
            if (_transition != null)
                _transition.Duration = value;
        }
    }

    /// <summary>
    /// 获取或设置速度
    /// </summary>
    public int Speed
    {
        get => _transition != null ? (int)_transition.Speed : 0;
        set
        {
            if (_transition != null)
                _transition.Speed = (MsPowerPoint.PpTransitionSpeed)value;
        }
    }



    /// <summary>
    /// 获取或设置是否循环
    /// </summary>
    public bool Loop
    {
        get => _transition?.LoopSoundUntilNext == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_transition != null)
                _transition.LoopSoundUntilNext = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="transition">COM SlideShowTransition 对象</param>
    internal PowerPointSlideShowTransition(MsPowerPoint.SlideShowTransition transition)
    {
        _transition = transition;
        _disposedValue = false;
    }


    /// <summary>
    /// 重置切换效果
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_transition != null)
            {
                _transition.EntryEffect = MsPowerPoint.PpEntryEffect.ppEffectNone;
                _transition.AdvanceOnTime = MsoTriState.msoCTrue;
                _transition.AdvanceTime = 0;
                _transition.Duration = 1;
                _transition.Speed = MsPowerPoint.PpTransitionSpeed.ppTransitionSpeedMedium;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset transition.", ex);
        }
    }

    /// <summary>
    /// 设置切换效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="duration">持续时间</param>
    /// <param name="speed">速度</param>
    public void SetTransition(int effectType, float duration = 1.0f, int speed = 2)
    {
        try
        {
            if (_transition != null)
            {
                _transition.EntryEffect = (MsPowerPoint.PpEntryEffect)effectType;
                _transition.Duration = duration;
                _transition.Speed = (MsPowerPoint.PpTransitionSpeed)speed;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set transition.", ex);
        }
    }

    /// <summary>
    /// 设置切换声音
    /// </summary>
    /// <param name="soundFile">声音文件路径</param>
    /// <param name="loop">是否循环</param>
    public void SetSound(string soundFile, bool loop = false)
    {
        try
        {
            if (_transition?.SoundEffect != null)
            {
                if (!string.IsNullOrEmpty(soundFile) && System.IO.File.Exists(soundFile))
                {
                    _transition.SoundEffect.ImportFromFile(soundFile);
                    _transition.LoopSoundUntilNext = loop ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set transition sound.", ex);
        }
    }

    /// <summary>
    /// 设置定时
    /// </summary>
    /// <param name="advanceTime">前进时间</param>
    /// <param name="advanceOnTime">是否定时前进</param>
    public void SetTiming(int advanceTime, bool advanceOnTime = true)
    {
        try
        {
            if (_transition != null)
            {
                _transition.AdvanceTime = advanceTime;
                _transition.AdvanceOnTime = advanceOnTime ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set transition timing.", ex);
        }
    }

    /// <summary>
    /// 应用到指定幻灯片范围
    /// </summary>
    /// <param name="fromSlide">起始幻灯片</param>
    /// <param name="toSlide">结束幻灯片</param>
    public void ApplyToRange(int fromSlide, int toSlide)
    {
        try
        {
            // 范围应用需要通过幻灯片集合实现
            throw new NotImplementedException("Applying transition to range is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply transition to range.", ex);
        }
    }

    /// <summary>
    /// 获取切换效果信息
    /// </summary>
    /// <returns>切换效果信息字符串</returns>
    public string GetTransitionInfo()
    {
        try
        {
            return $"Transition - Effect: {EntryEffect}, Duration: {Duration}, Speed: {Speed}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get transition info.", ex);
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

