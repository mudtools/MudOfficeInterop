//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// Office填充格式对象的实现类
/// </summary>
internal class OfficeFillFormat : IOfficeFillFormat
{
    private MsCore.FillFormat? _fillFormat;
    private bool _disposedValue;

    /// <summary>
    /// 初始化OfficeFillFormat类的新实例
    /// </summary>
    /// <param name="fillFormat">原始的COM填充格式对象</param>
    internal OfficeFillFormat(MsCore.FillFormat fillFormat)
    {
        _fillFormat = fillFormat ?? throw new ArgumentNullException(nameof(fillFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IOfficeColorFormat? ForeColor
    {
        get
        {
            if (_fillFormat == null)
                return null;

            var foreColor = _fillFormat.ForeColor;
            return foreColor != null ? new OfficeColorFormat(_fillFormat.ForeColor) : null;
        }
    }

    /// <inheritdoc/>
    public IOfficeColorFormat? BackColor
    {
        get
        {
            if (_fillFormat == null)
                return null;

            var backColor = _fillFormat.BackColor;
            return backColor != null ? new OfficeColorFormat(backColor) : null;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _fillFormat?.Transparency ?? 0;
        set
        {
            if (_fillFormat != null && value >= 0 && value <= 1)
                _fillFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public MsoFillType Type => _fillFormat?.Type != null ? _fillFormat.Type.EnumConvert(MsoFillType.msoFillMixed) : MsoFillType.msoFillMixed;

    /// <inheritdoc/>
    public float GradientAngle
    {
        get => _fillFormat?.GradientAngle ?? 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.GradientAngle = value;
        }
    }

    /// <inheritdoc/>
    public string TextureName => _fillFormat?.TextureName ?? string.Empty;

    /// <inheritdoc/>
    public MsoTextureType TextureType => _fillFormat?.TextureType != null ? _fillFormat.TextureType.EnumConvert(MsoTextureType.msoTextureTypeMixed) : MsoTextureType.msoTextureTypeMixed;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void OneColorGradient(MsoGradientStyle style, int variant, float degree)
    {
        try
        {
            if (_fillFormat != null)
                _fillFormat.OneColorGradient(style.EnumConvert(MsCore.MsoGradientStyle.msoGradientMixed), variant, degree);
        }
        catch (Exception ex)
        {
            throw new Exception("OneColorGradient方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void TwoColorGradient(MsoGradientStyle style, int variant)
    {
        try
        {
            if (_fillFormat != null)
                _fillFormat.TwoColorGradient(style.EnumConvert(MsCore.MsoGradientStyle.msoGradientMixed), variant);
        }
        catch (Exception ex)
        {
            throw new Exception("TwoColorGradient方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void Patterned(MsoPatternType pattern)
    {
        try
        {
            if (_fillFormat != null)
                _fillFormat.Patterned(pattern.EnumConvert(MsCore.MsoPatternType.msoPatternMixed));
        }
        catch (Exception ex)
        {
            throw new Exception("Patterned方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void UserPicture(string imagePath)
    {
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("图片路径不能为空！", nameof(imagePath));
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("图片文件不存在！", imagePath);
        try
        {
            if (_fillFormat != null && !string.IsNullOrWhiteSpace(imagePath))
            {
                _fillFormat.UserPicture(imagePath);
            }
        }
        catch (Exception ex)
        {
            throw new Exception("UserPicture方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void Solid()
    {
        try
        {
            _fillFormat?.Solid();
        }
        catch (Exception ex)
        {
            throw new Exception("Solid方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void UserTextured(string textureFile)
    {
        if (string.IsNullOrEmpty(textureFile))
            throw new ArgumentException("纹理文件路径不能为空！", nameof(textureFile));
        if (!File.Exists(textureFile))
            throw new FileNotFoundException("纹理文件不存在！", textureFile);
        try
        {
            if (_fillFormat != null && !string.IsNullOrWhiteSpace(textureFile))
            {
                _fillFormat.UserTextured(textureFile);
            }
        }
        catch (Exception ex)
        {
            throw new Exception("UserTextured方法调用失败！", ex);
        }
    }

    /// <inheritdoc/>
    public void Background()
    {
        try
        {
            _fillFormat?.Background();
        }
        catch (Exception ex)
        {
            throw new Exception("Background方法调用失败！", ex);
        }
    }

    #endregion

    #region IDisposable实现

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在处置</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fillFormat != null)
        {
            Marshal.ReleaseComObject(_fillFormat);
            _fillFormat = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}