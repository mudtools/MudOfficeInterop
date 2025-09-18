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