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
    private MsCore.FillFormat _fillFormat;
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
            return foreColor != null ? new OfficeColorFormat(foreColor) : null;
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
    public MsoFillType Type => _fillFormat?.Type != null ? (MsoFillType)(int)_fillFormat?.Type : MsoFillType.msoFillMixed;

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
    public MsoTextureType TextureType => _fillFormat?.TextureType != null ? (MsoTextureType)(int)_fillFormat?.TextureType : MsoTextureType.msoTextureTypeMixed;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void OneColorGradient(MsoGradientStyle style, int variant, float degree)
    {
        _fillFormat?.OneColorGradient((MsCore.MsoGradientStyle)(int)style, variant, degree);
    }

    /// <inheritdoc/>
    public void TwoColorGradient(MsoGradientStyle style, int variant)
    {
        _fillFormat?.TwoColorGradient((MsCore.MsoGradientStyle)(int)style, variant);
    }

    /// <inheritdoc/>
    public void Patterned(MsoPatternType pattern)
    {
        _fillFormat?.Patterned((MsCore.MsoPatternType)(int)pattern);
    }

    /// <inheritdoc/>
    public void UserPicture(string imagePath)
    {
        if (_fillFormat != null && !string.IsNullOrWhiteSpace(imagePath))
        {
            _fillFormat.UserPicture(imagePath);
        }
    }

    /// <inheritdoc/>
    public void Solid()
    {
        _fillFormat?.Solid();
    }

    /// <inheritdoc/>
    public void UserTextured(string textureFile)
    {
        if (_fillFormat != null && !string.IsNullOrWhiteSpace(textureFile))
        {
            _fillFormat?.UserTextured(textureFile);
        }
    }

    /// <inheritdoc/>
    public void Background()
    {
        _fillFormat?.Background();
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