//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
            try
            {
                if (_fillFormat != null)
                    Marshal.ReleaseComObject(_fillFormat);
            }
            catch { }
            _fillFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public MsoFillType Type
    {
        get => _fillFormat != null ? (MsoFillType)_fillFormat.Type : MsoFillType.msoFillSolid;
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
        get => _fillFormat != null ? (MsoPatternType)_fillFormat.Pattern : MsoPatternType.msoPattern5Percent;
    }

    public int Transparency
    {
        get => _fillFormat != null ? Convert.ToInt32(_fillFormat.Transparency * 100) : 0;
        set
        {
            if (_fillFormat != null)
                _fillFormat.Transparency = value / 100.0f;
        }
    }

    public bool Visible
    {
        get => _fillFormat != null && Convert.ToBoolean(_fillFormat.Visible);
        set
        {
            if (_fillFormat != null)
                _fillFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    public MsoGradientColorType GradientColorType
    {
        get
        {
            return (MsoGradientColorType)_fillFormat.GradientColorType;
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
            return (MsoGradientStyle)_fillFormat.GradientStyle;
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
            return (MsoPresetGradientType)_fillFormat.PresetGradientType;
        }
    }
    public MsoPresetTexture PresetTexture
    {
        get
        {
            return (MsoPresetTexture)_fillFormat.PresetTexture;
        }
    }

    public string TextureName
    {
        get
        {
            return _fillFormat.TextureName;
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
}
