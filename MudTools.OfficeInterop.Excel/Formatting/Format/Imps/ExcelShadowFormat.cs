//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ShadowFormat 对象的二次封装实现类
/// </summary>
internal class ExcelShadowFormat : IExcelShadowFormat
{
    private MsExcel.ShadowFormat _shadowFormat;
    private bool _disposedValue;

    internal ExcelShadowFormat(MsExcel.ShadowFormat shadowFormat)
    {
        _shadowFormat = shadowFormat;
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_shadowFormat != null)
                Marshal.ReleaseComObject(_shadowFormat);
            _shadowFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public MsoShadowType Type
    {
        get => _shadowFormat != null ? _shadowFormat.Type.EnumConvert(MsoShadowType.msoShadowMixed) : MsoShadowType.msoShadowMixed;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Type = value.EnumConvert(MsCore.MsoShadowType.msoShadowMixed);
        }
    }

    public MsoShadowStyle Style
    {
        get => _shadowFormat != null ? _shadowFormat.Style.EnumConvert(MsoShadowStyle.msoShadowStyleMixed) : MsoShadowStyle.msoShadowStyleMixed;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Style = value.EnumConvert(MsCore.MsoShadowStyle.msoShadowStyleMixed);
        }
    }

    public IExcelColorFormat? ForeColor
    {
        get => _shadowFormat != null ? new ExcelColorFormat(_shadowFormat.ForeColor) : null;
        set
        {
            if (_shadowFormat != null && value is ExcelColorFormat format)
            {
                _shadowFormat.ForeColor = format.InternalComObject;
            }
        }
    }

    public int Color
    {
        get => _shadowFormat != null ? Convert.ToInt32(_shadowFormat.ForeColor.RGB) : 0;
        set { if (_shadowFormat != null) _shadowFormat.ForeColor.RGB = value; }
    }

    public int Transparency
    {
        get => _shadowFormat != null ? Convert.ToInt32(_shadowFormat.Transparency * 100) : 0;
        set { if (_shadowFormat != null) _shadowFormat.Transparency = value / 100.0f; }
    }

    public float OffsetX
    {
        get => _shadowFormat?.OffsetX ?? 0;
        set { if (_shadowFormat != null) _shadowFormat.OffsetX = value; }
    }

    public float OffsetY
    {
        get => _shadowFormat?.OffsetY ?? 0;
        set { if (_shadowFormat != null) _shadowFormat.OffsetY = value; }
    }

    public float Blur
    {
        get => _shadowFormat?.Blur ?? 0;
        set { if (_shadowFormat != null) _shadowFormat.Blur = value; }
    }

    public float Size
    {
        get => _shadowFormat?.Size ?? 0;
        set { if (_shadowFormat != null) _shadowFormat.Size = value; }
    }

    public bool Visible
    {
        get => _shadowFormat != null && _shadowFormat.Visible.ConvertToBool();
        set { if (_shadowFormat != null) _shadowFormat.Visible = value.ConvertTriState(); }
    }

    public bool Obscured
    {
        get => _shadowFormat != null && _shadowFormat.Obscured.ConvertToBool();
        set { if (_shadowFormat != null) _shadowFormat.Obscured = value.ConvertTriState(); }
    }

    public bool RotateWithShape
    {
        get => _shadowFormat != null && _shadowFormat.RotateWithShape.ConvertToBool();
        set { if (_shadowFormat != null) _shadowFormat.RotateWithShape = value.ConvertTriState(); }
    }

    public void IncrementOffsetX(float Increment)
    {
        _shadowFormat?.IncrementOffsetX(Increment);
    }

    public void IncrementOffsetY(float Increment)
    {
        _shadowFormat?.IncrementOffsetY(Increment);
    }
}