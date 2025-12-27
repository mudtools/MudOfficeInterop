//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 阴影格式实现类
/// </summary>
internal class PowerPointShadowFormat : IPowerPointShadowFormat
{
    private readonly MsPowerPoint.ShadowFormat _shadowFormat;
    private bool _disposedValue;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _shadowFormat?.Parent;

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    public bool Visible
    {
        get => _shadowFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置阴影类型
    /// </summary>
    public int Type
    {
        get => _shadowFormat != null ? (int)_shadowFormat.Type : 0;
        set
        {
            // Type 是只读属性
        }
    }

    /// <summary>
    /// 获取或设置水平偏移
    /// </summary>
    public float OffsetX
    {
        get => _shadowFormat?.OffsetX ?? 0;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetX = value;
        }
    }

    /// <summary>
    /// 获取或设置垂直偏移
    /// </summary>
    public float OffsetY
    {
        get => _shadowFormat?.OffsetY ?? 0;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetY = value;
        }
    }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    public int ForeColor
    {
        get => _shadowFormat?.ForeColor.RGB ?? 0;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.ForeColor.RGB = value;
        }
    }

    /// <summary>
    /// 获取或设置模糊度
    /// </summary>
    public float Blur
    {
        get => _shadowFormat?.Blur ?? 0;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Blur = value;
        }
    }

    /// <summary>
    /// 获取或设置阴影大小
    /// </summary>
    public float Size
    {
        get => _shadowFormat?.Size ?? 0;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Size = value;
        }
    }


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="shadowFormat">COM ShadowFormat 对象</param>
    internal PowerPointShadowFormat(MsPowerPoint.ShadowFormat shadowFormat)
    {
        _shadowFormat = shadowFormat; // 可以为 null
        _disposedValue = false;
    }


    /// <summary>
    /// 重置阴影格式
    /// </summary>
    public void Reset()
    {
        try
        {
            if (_shadowFormat != null) _shadowFormat.Visible = MsCore.MsoTriState.msoFalse;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset shadow format.", ex);
        }
    }

    /// <summary>
    /// 复制阴影格式
    /// </summary>
    /// <returns>复制的阴影格式对象</returns>
    public IPowerPointShadowFormat Duplicate()
    {
        try
        {
            // PowerPoint 中没有直接的复制方法
            throw new NotImplementedException("Duplicating shadow format is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate shadow format.", ex);
        }
    }

    /// <summary>
    /// 应用阴影格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    public void ApplyTo(IPowerPointShape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        try
        {
            // 这需要具体的实现来应用阴影格式到形状
            throw new NotImplementedException("Applying shadow format to shape is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply shadow format to shape.", ex);
        }
    }

    /// <summary>
    /// 设置阴影偏移
    /// </summary>
    /// <param name="offsetX">水平偏移</param>
    /// <param name="offsetY">垂直偏移</param>
    public void SetOffset(float offsetX, float offsetY)
    {
        try
        {
            if (_shadowFormat != null)
            {
                _shadowFormat.OffsetX = offsetX;
                _shadowFormat.OffsetY = offsetY;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set shadow offset.", ex);
        }
    }

    /// <summary>
    /// 设置阴影颜色
    /// </summary>
    /// <param name="color">阴影颜色</param>
    /// <param name="transparency">透明度</param>
    public void SetColor(int color, float transparency = 0)
    {
        try
        {
            if (_shadowFormat != null)
            {
                _shadowFormat.ForeColor.RGB = color;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set shadow color.", ex);
        }
    }

    /// <summary>
    /// 设置阴影大小和模糊度
    /// </summary>
    /// <param name="size">阴影大小</param>
    /// <param name="blur">模糊度</param>
    public void SetSizeAndBlur(float size, float blur)
    {
        try
        {
            if (_shadowFormat != null)
            {
                _shadowFormat.Size = size;
                _shadowFormat.Blur = blur;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set shadow size and blur.", ex);
        }
    }



    /// <summary>
    /// 获取阴影信息
    /// </summary>
    /// <returns>阴影信息字符串</returns>
    public string GetShadowInfo()
    {
        try
        {
            return $"Shadow Visible: {Visible}, OffsetX: {OffsetX}, OffsetY: {OffsetY}, Color: {ForeColor}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get shadow info.", ex);
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
