//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.FillFormat 的实现类。
/// </summary>
internal class WordFillFormat : IWordFillFormat
{
    private MsWord.FillFormat _fillFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="fillFormat">原始 COM FillFormat 对象。</param>
    internal WordFillFormat(MsWord.FillFormat fillFormat)
    {
        _fillFormat = fillFormat ?? throw new ArgumentNullException(nameof(fillFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _fillFormat != null ? new WordApplication(_fillFormat.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _fillFormat?.Parent;

    /// <inheritdoc/>
    public IWordColorFormat? ForeColor =>
        _fillFormat?.ForeColor != null ? new WordColorFormat(_fillFormat.ForeColor) : null;

    /// <inheritdoc/>
    public IWordColorFormat? BackColor =>
        _fillFormat?.BackColor != null ? new WordColorFormat(_fillFormat.BackColor) : null;

    /// <inheritdoc/>
    public float Transparency
    {
        get => _fillFormat?.Transparency ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public bool Visible
    {
        get => _fillFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_fillFormat != null)
                _fillFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public MsoFillType Type => _fillFormat?.Type != null ? (MsoFillType)(int)_fillFormat?.Type : MsoFillType.msoFillSolid;

    /// <inheritdoc/>
    public MsoGradientColorType GradientColorType
    {
        get => _fillFormat?.GradientColorType != null ? (MsoGradientColorType)(int)_fillFormat?.GradientColorType : MsoGradientColorType.msoGradientColorMixed;

    }

    /// <inheritdoc/>
    public MsoGradientStyle GradientStyle
    {
        get => _fillFormat?.GradientStyle != null ? (MsoGradientStyle)(int)_fillFormat?.GradientStyle : MsoGradientStyle.msoGradientHorizontal;
    }

    /// <inheritdoc/>
    public float GradientAngle
    {
        get => _fillFormat?.GradientAngle ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.GradientAngle = value;
        }
    }

    /// <inheritdoc/>
    public MsoPatternType Pattern
    {
        get => _fillFormat?.Pattern != null ? (MsoPatternType)(int)_fillFormat?.Pattern : MsoPatternType.msoPatternMixed;
    }

    /// <inheritdoc/>
    public MsoPresetTexture PresetTexture
    {
        get => _fillFormat?.PresetTexture != null ? (MsoPresetTexture)(int)_fillFormat?.PresetTexture : MsoPresetTexture.msoPresetTextureMixed;
    }

    /// <inheritdoc/>
    public string TextureName
    {
        get => _fillFormat?.TextureName ?? string.Empty;
    }

    /// <inheritdoc/>
    public MsoTextureType TextureType
    {
        get => _fillFormat?.TextureType != null ? (MsoTextureType)(int)_fillFormat?.TextureType : MsoTextureType.msoTextureTypeMixed;
    }

    /// <inheritdoc/>
    public MsoPresetGradientType PresetGradientType
    {
        get => _fillFormat?.PresetGradientType != null ? (MsoPresetGradientType)(int)_fillFormat?.PresetGradientType : MsoPresetGradientType.msoPresetGradientMixed;
    }

    /// <inheritdoc/>
    public int GradientStopsCount => _fillFormat?.GradientStops?.Count ?? 0;


    public float TextureOffsetX
    {
        get => _fillFormat?.TextureOffsetX ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureOffsetX = value;
        }
    }

    public float TextureOffsetY
    {
        get => _fillFormat?.TextureOffsetY ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureOffsetY = value;
        }
    }

    public MsoTextureAlignment TextureAlignment
    {
        get => _fillFormat?.TextureAlignment != null ? (MsoTextureAlignment)(int)_fillFormat?.TextureAlignment : MsoTextureAlignment.msoTextureLeft;
        set => _fillFormat.TextureAlignment = (MsCore.MsoTextureAlignment)(int)value;
    }

    public float TextureHorizontalScale
    {
        get => _fillFormat?.TextureHorizontalScale ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureHorizontalScale = value;
        }
    }

    public float TextureVerticalScale
    {
        get => _fillFormat?.TextureVerticalScale ?? 0f;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureVerticalScale = value;
        }
    }

    public bool TextureTile
    {
        get => _fillFormat?.TextureTile == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureTile = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    public bool RotateWithObject
    {
        get => _fillFormat?.RotateWithObject == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_fillFormat != null)
                _fillFormat.RotateWithObject = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    public IOfficePictureEffects PictureEffects
    {
        get => _fillFormat?.PictureEffects != null ? new OfficePictureEffects(_fillFormat.PictureEffects) : null;
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Solid()
    {
        _fillFormat?.Solid();
    }

    /// <inheritdoc/>
    public void Solid(int color)
    {
        if (_fillFormat != null)
        {
            _fillFormat.Solid();
            _fillFormat.ForeColor.RGB = color;
        }
    }

    /// <inheritdoc/>
    public void PresetGradient(MsoGradientStyle style, int variant, MsoPresetGradientType presetGradientType)
    {
        _fillFormat?.PresetGradient((MsCore.MsoGradientStyle)(int)style, variant, (MsCore.MsoPresetGradientType)(int)presetGradientType);
    }

    /// <inheritdoc/>
    public void PresetTextured(MsoPresetTexture presetTexture)
    {
        _fillFormat?.PresetTextured((MsCore.MsoPresetTexture)(int)presetTexture);
    }

    /// <inheritdoc/>
    public void Patterned(MsoPatternType pattern, int foregroundColor, int backgroundColor)
    {
        if (_fillFormat != null)
        {
            _fillFormat.Patterned((MsCore.MsoPatternType)(int)pattern);
            if (_fillFormat.ForeColor != null)
                _fillFormat.ForeColor.RGB = foregroundColor;
            if (_fillFormat.BackColor != null)
                _fillFormat.BackColor.RGB = backgroundColor;
        }
    }

    /// <inheritdoc/>
    public void UserPicture(string pictureFile)
    {
        if (_fillFormat != null && !string.IsNullOrWhiteSpace(pictureFile))
        {
            _fillFormat.UserPicture(pictureFile);
        }
    }

    /// <inheritdoc/>
    public void UserTextured(string textureFile)
    {
        if (_fillFormat != null && !string.IsNullOrWhiteSpace(textureFile))
        {
            _fillFormat.UserTextured(textureFile);
        }
    }

    /// <inheritdoc/>
    public void SetTransparent(float transparency)
    {
        if (_fillFormat != null)
        {
            _fillFormat.Transparency = transparency;
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_fillFormat != null)
        {
            _fillFormat.Visible = 0;
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordFillFormat targetFill)
    {
        if (_fillFormat == null || targetFill == null)
            return;

        try
        {
            targetFill.Transparency = this.Transparency;
            targetFill.Visible = this.Visible;

            // 复制颜色属性
            if (this.ForeColor != null && targetFill.ForeColor != null)
            {
                targetFill.ForeColor.RGB = this.ForeColor.RGB;
            }
            if (this.BackColor != null && targetFill.BackColor != null)
            {
                targetFill.BackColor.RGB = this.BackColor.RGB;
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制填充格式。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放前景色对象
            if (_fillFormat?.ForeColor != null)
            {
                Marshal.ReleaseComObject(_fillFormat.ForeColor);
            }
            // 释放背景色对象
            if (_fillFormat?.BackColor != null)
            {
                Marshal.ReleaseComObject(_fillFormat.BackColor);
            }
            // 释放渐变停靠点集合
            if (_fillFormat?.GradientStops != null)
            {
                Marshal.ReleaseComObject(_fillFormat.GradientStops);
            }
            // 释放填充格式对象本身
            if (_fillFormat != null)
            {
                Marshal.ReleaseComObject(_fillFormat);
                _fillFormat = null;
            }
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