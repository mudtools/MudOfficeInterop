
namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// ReflectionFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class OfficeReflectionFormat : IOfficeReflectionFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsCore.ReflectionFormat _reflectionFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="reflectionFormat">原始的 ReflectionFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 reflectionFormat 为 null 时抛出。</exception>
    internal OfficeReflectionFormat(MsCore.ReflectionFormat reflectionFormat)
    {
        _reflectionFormat = reflectionFormat ?? throw new ArgumentNullException(nameof(reflectionFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_reflectionFormat != null)
            {
                Marshal.ReleaseComObject(_reflectionFormat);
                _reflectionFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);


    /// <summary>
    /// 获取或设置倒影的类型（无倒影、预设倒影样式等）。
    /// 默认值：msoReflectionTypeMixed
    /// 设置时自动转换为底层 COM 枚举类型。
    /// </summary>
    public MsoReflectionType Type
    {
        get => _reflectionFormat != null
            ? _reflectionFormat.Type.EnumConvert(MsoReflectionType.msoReflectionTypeMixed)
            : MsoReflectionType.msoReflectionTypeMixed;

        set
        {
            if (_reflectionFormat != null)
            {
                _reflectionFormat.Type = value.EnumConvert(MsCore.MsoReflectionType.msoReflectionTypeMixed);
            }
        }
    }

    /// <summary>
    /// 获取或设置倒影的透明度（0-100，0=完全不透明，100=完全透明）。
    /// 内部自动转换为 COM 所需的 0.0~1.0 浮点值。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public int Transparency
    {
        get => _reflectionFormat != null ? Convert.ToInt32(_reflectionFormat.Transparency * 100) : 0;

        set
        {
            if (_reflectionFormat != null)
            {
                float val = Math.Max(0, Math.Min(100, value)) / 100.0f;
                _reflectionFormat.Transparency = val;
            }
        }
    }

    /// <summary>
    /// 获取或设置倒影的大小比例（0.0~1.0，1.0=100% 原图高度）。
    /// 例如：0.5 表示倒影高度为原对象的一半。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public float Size
    {
        get => _reflectionFormat != null ? _reflectionFormat.Size : 0f;

        set
        {
            if (_reflectionFormat != null)
            {
                if (value < 0f || value > 1.0f)
                    throw new ArgumentOutOfRangeException(nameof(value), "Size 必须在 0.0 到 1.0 之间。");
                _reflectionFormat.Size = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置倒影的模糊程度（单位：磅）。
    /// 值越大，倒影边缘越模糊。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public float Blur
    {
        get => _reflectionFormat != null ? _reflectionFormat.Blur : 0f;

        set
        {
            if (_reflectionFormat != null)
            {
                if (value < 0f)
                    throw new ArgumentOutOfRangeException(nameof(value), "Blur 不能小于 0。");
                _reflectionFormat.Blur = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置倒影与原对象的垂直距离（单位：磅）。
    /// 正值表示倒影在对象下方。
    /// 若 COM 对象为空，设置无效，获取返回 0。
    /// </summary>
    public float Offset
    {
        get => _reflectionFormat != null ? _reflectionFormat.Offset : 0f;

        set
        {
            if (_reflectionFormat != null)
            {
                _reflectionFormat.Offset = value;
            }
        }
    }

    /// <summary>
    /// 获取倒影效果是否已启用（Type != msoReflectionTypeNone）。
    /// </summary>
    public bool Visible =>
        _reflectionFormat != null &&
        _reflectionFormat.Type.EnumConvert(MsoReflectionType.msoReflectionTypeNone) != MsoReflectionType.msoReflectionTypeNone;
}