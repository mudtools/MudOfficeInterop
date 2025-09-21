//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelPictureFormat : IExcelPictureFormat
{
    internal MsExcel.PictureFormat _pictureFormat;
    private bool _disposedValue;

    internal ExcelPictureFormat(MsExcel.PictureFormat pictureFormat)
    {
        _pictureFormat = pictureFormat ?? throw new ArgumentNullException(nameof(pictureFormat));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_pictureFormat != null)
                Marshal.ReleaseComObject(_pictureFormat);
            _pictureFormat = null;
        }
        _disposedValue = true;
    }

    public void Dispose() => Dispose(true);

    public object Parent => _pictureFormat?.Parent;

    public IExcelApplication Application =>
        _pictureFormat?.Application != null ? new ExcelApplication(_pictureFormat.Application as MsExcel.Application) : null;

    public float Brightness
    {
        get => _pictureFormat != null ? _pictureFormat.Brightness : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Brightness = value;
        }
    }

    public MsoPictureColorType ColorType
    {
        get => _pictureFormat != null ? _pictureFormat.ColorType.EnumConvert(MsoPictureColorType.msoPictureAutomatic) : MsoPictureColorType.msoPictureAutomatic;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.ColorType = value.EnumConvert(MsCore.MsoPictureColorType.msoPictureAutomatic);
        }
    }

    public float Contrast
    {
        get => _pictureFormat != null ? _pictureFormat.Contrast : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Contrast = value;
        }
    }

    public float CropBottom
    {
        get => _pictureFormat != null ? _pictureFormat.CropBottom : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropBottom = value;
        }
    }

    public float CropLeft
    {
        get => _pictureFormat != null ? _pictureFormat.CropLeft : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropLeft = value;
        }
    }

    public float CropRight
    {
        get => _pictureFormat != null ? _pictureFormat.CropRight : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropRight = value;
        }
    }

    public float CropTop
    {
        get => _pictureFormat != null ? _pictureFormat.CropTop : 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropTop = value;
        }
    }

    public bool TransparentBackground
    {
        get => _pictureFormat != null && _pictureFormat.TransparentBackground.ConvertToBool();
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.TransparentBackground = value.ConvertTriState();
        }
    }

    public void IncrementBrightness(float Increment)
    {
        _pictureFormat?.IncrementBrightness(Increment);
    }

    public void IncrementContrast(float Increment)
    {
        _pictureFormat?.IncrementContrast(Increment);
    }
}