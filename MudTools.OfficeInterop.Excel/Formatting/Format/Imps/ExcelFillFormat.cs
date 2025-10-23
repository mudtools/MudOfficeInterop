//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel FillFormat 对象的二次封装实现类
/// </summary>
internal class ExcelFillFormat : IExcelFillFormat
{
    internal MsExcel.FillFormat _fillFormat;
    private bool _disposedValue;

    internal ExcelFillFormat(MsExcel.FillFormat fillFormat)
    {
        _fillFormat = fillFormat;
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_fillFormat != null)
                Marshal.ReleaseComObject(_fillFormat);
            _fillFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object Parent => _fillFormat.Parent;

    public IExcelApplication Application => new ExcelApplication(_fillFormat.Application as MsExcel.Application);

    public MsoFillType Type
    {
        get => _fillFormat != null ? _fillFormat.Type.EnumConvert(MsoFillType.msoFillSolid) : MsoFillType.msoFillSolid;
    }

    public IExcelColorFormat ForeColor
    {
        get => new ExcelColorFormat(_fillFormat.ForeColor);
        set => _fillFormat.ForeColor = ((ExcelColorFormat)value)._colorFormat;
    }

    public IExcelColorFormat BackColor
    {
        get => new ExcelColorFormat(_fillFormat.BackColor);
        set => _fillFormat.BackColor = ((ExcelColorFormat)value)._colorFormat;
    }

    public MsoPatternType Pattern
    {
        get => _fillFormat != null ? _fillFormat.Pattern.EnumConvert(MsoPatternType.msoPattern5Percent) : MsoPatternType.msoPattern5Percent;
    }

    public float Transparency
    {
        get => _fillFormat != null ? _fillFormat.Transparency : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.Transparency = value;
        }
    }

    public bool Visible
    {
        get => _fillFormat != null && _fillFormat.Visible.ConvertToBool();
        set
        {
            if (_fillFormat != null)
                _fillFormat.Visible = value.ConvertTriState();
        }
    }

    public MsoGradientColorType GradientColorType
    {
        get
        {
            return _fillFormat.GradientColorType.EnumConvert(MsoGradientColorType.msoGradientColorMixed);
        }
    }

    public float GradientDegree
    {
        get
        {
            return _fillFormat.GradientDegree;
        }
    }

    public MsoGradientStyle GradientStyle
    {
        get
        {
            return _fillFormat.GradientStyle.EnumConvert(MsoGradientStyle.msoGradientMixed);
        }
    }


    public int GradientVariant
    {
        get
        {
            return _fillFormat.GradientVariant;
        }
    }

    public MsoPresetGradientType PresetGradientType
    {
        get
        {
            return _fillFormat.PresetGradientType.EnumConvert(MsoPresetGradientType.msoPresetGradientMixed);
        }
    }
    public MsoPresetTexture PresetTexture
    {
        get
        {
            return _fillFormat.PresetTexture.EnumConvert(MsoPresetTexture.msoPresetTextureMixed);
        }
    }

    public MsoTextureType TextureType
    {
        get
        {
            return _fillFormat.TextureType.EnumConvert(MsoTextureType.msoTextureTypeMixed);
        }
    }

    public MsoTextureAlignment TextureAlignment
    {
        get
        {
            return _fillFormat.TextureAlignment.EnumConvert(MsoTextureAlignment.msoTextureAlignmentMixed);
        }
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureAlignment = value.EnumConvert(MsCore.MsoTextureAlignment.msoTextureAlignmentMixed);
        }
    }

    public string TextureName
    {
        get
        {
            return _fillFormat.TextureName;
        }
    }

    public float TextureOffsetX
    {
        get => _fillFormat != null ? _fillFormat.TextureOffsetX : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureOffsetX = value;
        }
    }
    public float TextureOffsetY
    {
        get => _fillFormat != null ? _fillFormat.TextureOffsetY : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureOffsetY = value;
        }
    }

    public float TextureHorizontalScale
    {
        get => _fillFormat != null ? _fillFormat.TextureHorizontalScale : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureHorizontalScale = value;
        }
    }
    public float TextureVerticalScale
    {
        get => _fillFormat != null ? _fillFormat.TextureVerticalScale : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureVerticalScale = value;
        }
    }

    public bool TextureTile
    {
        get => _fillFormat != null ? _fillFormat.TextureTile.ConvertToBool() : false;
        set
        {
            if (_fillFormat != null)
                _fillFormat.TextureTile = value.ConvertTriState();
        }
    }


    public void UserPicture(string PictureFile)
    {
        _fillFormat?.UserPicture(PictureFile);
    }

    public void UserTextured(string TextureFile)
    {
        _fillFormat?.UserTextured(TextureFile);
    }

    public void OneColorGradient(MsoGradientStyle style, int variant, float degree)
    {
        _fillFormat?.OneColorGradient(style.EnumConvert(MsCore.MsoGradientStyle.msoGradientMixed), variant, degree);
    }

    public void Patterned(MsoPatternType pattern)
    {
        _fillFormat?.Patterned(pattern.EnumConvert(MsCore.MsoPatternType.msoPatternMixed));
    }

    public void PresetTextured(MsoPresetTexture PresetTexture)
    {
        _fillFormat?.PresetTextured(PresetTexture.EnumConvert(MsCore.MsoPresetTexture.msoPresetTextureMixed));
    }


    public void TwoColorGradient(MsoGradientStyle style, int variant)
    {
        _fillFormat?.TwoColorGradient(style.EnumConvert(MsCore.MsoGradientStyle.msoGradientMixed), variant);
    }

    public void Solid()
    {
        _fillFormat?.Solid();
    }

    public void Background()
    {
        _fillFormat?.Background();
    }
}
