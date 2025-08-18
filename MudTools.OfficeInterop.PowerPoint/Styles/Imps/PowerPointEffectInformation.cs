//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 效果信息实现类
/// </summary>
internal class PowerPointEffectInformation : IPowerPointEffectInformation
{
    private readonly MsPowerPoint.Effect _effect;
    private bool _disposedValue;

    /// <summary>
    /// 获取显示名称
    /// </summary>
    public string DisplayName
    {
        get
        {
            try
            {
                return _effect?.DisplayName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// 获取效果类型
    /// </summary>
    public int EffectType => _effect != null ? (int)_effect.EffectType : 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _effect?.Parent;



    /// <summary>
    /// 获取或设置触发延迟
    /// </summary>
    public float TriggerDelayTime
    {
        get => _effect?.Timing?.TriggerDelayTime ?? 0;
        set
        {
            if (_effect?.Timing != null)
                _effect.Timing.TriggerDelayTime = value;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="effect">COM Effect 对象</param>
    internal PowerPointEffectInformation(MsPowerPoint.Effect effect)
    {
        _effect = effect;
        _disposedValue = false;
    }

    /// <summary>
    /// 获取效果信息
    /// </summary>
    /// <returns>效果信息字符串</returns>
    public string GetEffectInformation()
    {
        try
        {
            return $"EffectInformation - Type: {EffectType}, DisplayName: {DisplayName}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get effect information.", ex);
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

