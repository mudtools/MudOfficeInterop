//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 动画效果实现类
/// </summary>
internal class PowerPointEffect : IPowerPointEffect
{
    private readonly MsPowerPoint.Effect _effect;
    private bool _disposedValue;
    private IPowerPointShape _shape;
    private IPowerPointEffectInformation _effectInformation;

    /// <summary>
    /// 获取目标形状
    /// </summary>
    public IPowerPointShape Shape
    {
        get
        {
            if (_shape == null && _effect?.Shape != null)
            {
                _shape = new PowerPointShape(_effect.Shape);
            }
            return _shape;
        }
    }

    /// <summary>
    /// 获取或设置效果类型
    /// </summary>
    public int EffectType
    {
        get => _effect != null ? (int)_effect.EffectType : 0;
        set
        {
            if (_effect != null)
                _effect.EffectType = (MsPowerPoint.MsoAnimEffect)value;
        }
    }

    /// <summary>
    /// 获取效果信息
    /// </summary>
    public IPowerPointEffectInformation EffectInformation
    {
        get
        {
            if (_effectInformation == null && _effect != null)
            {
                _effectInformation = new PowerPointEffectInformation(_effect);
            }
            return _effectInformation;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _effect?.Parent;

    /// <summary>
    /// 获取或设置效果索引
    /// </summary>
    public int Index
    {
        get => _effect?.Index ?? 0;
        set
        {
            // Index 是只读属性
        }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="effect">COM Effect 对象</param>
    internal PowerPointEffect(MsPowerPoint.Effect effect)
    {
        _effect = effect;
        _disposedValue = false;
    }

    /// <summary>
    /// 应用效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="triggerType">触发类型</param>
    public void ApplyEffect(int effectType, int triggerType = 1)
    {
        try
        {
            if (_effect != null)
            {
                _effect.EffectType = (MsPowerPoint.MsoAnimEffect)effectType;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply effect.", ex);
        }
    }

    /// <summary>
    /// 删除效果
    /// </summary>
    public void Delete()
    {
        try
        {
            _effect?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete effect.", ex);
        }
    }

    /// <summary>
    /// 移动效果到指定位置
    /// </summary>
    /// <param name="index">目标位置</param>
    public void MoveTo(int index)
    {
        try
        {
            _effect?.MoveTo(index);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to move effect to index {index}.", ex);
        }
    }

    /// <summary>
    /// 设置效果参数
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <param name="value">属性值</param>
    public void SetProperty(string propertyName, object value)
    {
        if (string.IsNullOrEmpty(propertyName))
            throw new ArgumentException("Property name cannot be null or empty.", nameof(propertyName));

        try
        {
            // 使用反射设置属性值
            throw new NotImplementedException("Setting effect property is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to set property '{propertyName}'.", ex);
        }
    }

    /// <summary>
    /// 获取效果参数
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <returns>属性值</returns>
    public object GetProperty(string propertyName)
    {
        if (string.IsNullOrEmpty(propertyName))
            throw new ArgumentException("Property name cannot be null or empty.", nameof(propertyName));

        try
        {
            // 使用反射获取属性值
            throw new NotImplementedException("Getting effect property is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get property '{propertyName}'.", ex);
        }
    }

    /// <summary>
    /// 预览效果
    /// </summary>
    public void Preview()
    {
        try
        {
            // 效果预览需要通过幻灯片放映实现
            throw new NotImplementedException("Previewing effect is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to preview effect.", ex);
        }
    }

    /// <summary>
    /// 获取效果信息
    /// </summary>
    /// <returns>效果信息字符串</returns>
    public string GetEffectInfo()
    {
        try
        {
            return $"Effect - Type: {EffectType}, Shape: {Shape?.Name}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get effect info.", ex);
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
            _shape?.Dispose();
            _effectInformation?.Dispose();
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
