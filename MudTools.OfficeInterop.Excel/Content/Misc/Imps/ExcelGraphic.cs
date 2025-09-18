//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelGraphic : IExcelGraphic
{
    private MsExcel.Graphic _graphic;
    private bool _disposedValue;

    public object Parent => _graphic.Parent;


    public float Width
    {
        get => _graphic.Width;
        set => _graphic.Width = value;
    }

    public float Height
    {
        get => _graphic.Height;
        set => _graphic.Height = value;
    }

    public bool LockAspectRatio
    {
        get => _graphic.LockAspectRatio == MsCore.MsoTriState.msoTrue;
        set => _graphic.LockAspectRatio = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
    }

    public float Brightness
    {
        get => _graphic.Brightness;
        set => _graphic.Brightness = value;
    }

    public float Contrast
    {
        get => _graphic.Contrast;
        set => _graphic.Contrast = value;
    }

    public MsoPictureColorType ColorType
    {
        get => _graphic.ColorType.EnumConvert(MsoPictureColorType.msoPictureAutomatic);
        set => _graphic.ColorType = value.EnumConvert(MsCore.MsoPictureColorType.msoPictureAutomatic);
    }

    public string Filename
    {
        get => _graphic.Filename;
        set => _graphic.Filename = value;
    }


    public bool IsCropped => _graphic.CropLeft != 0 || _graphic.CropRight != 0 ||
                            _graphic.CropTop != 0 || _graphic.CropBottom != 0;

    public float CropLeft
    {
        get => _graphic.CropLeft;
        set => _graphic.CropLeft = value;
    }

    public float CropRight
    {
        get => _graphic.CropRight;
        set => _graphic.CropRight = value;
    }

    public float CropTop
    {
        get => _graphic.CropTop;
        set => _graphic.CropTop = value;
    }

    public float CropBottom
    {
        get => _graphic.CropBottom;
        set => _graphic.CropBottom = value;
    }

    internal ExcelGraphic(MsExcel.Graphic graphic)
    {
        _graphic = graphic ?? throw new ArgumentNullException(nameof(graphic));
        _disposedValue = false;
    }
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _graphic != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_graphic) > 0) { }
            }
            catch { }
            _graphic = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}