//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Picture 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Picture 对象的安全访问和资源管理
/// </summary>
internal class ExcelPicture : IExcelPicture
{
    /// <summary>
    /// 底层的 COM Picture 对象
    /// </summary>
    private MsExcel.Picture _picture;


    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelPicture 实例
    /// </summary>
    /// <param name="picture">底层的 COM Picture 对象</param>
    internal ExcelPicture(MsExcel.Picture picture)
    {
        _picture = picture ?? throw new ArgumentNullException(nameof(picture));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _bottomRightCellRange?.Dispose();
            _topLeftCellRange?.Dispose();
            _interior?.Dispose();
            _excelShape?.Dispose();
            _border?.Dispose();

            // 释放底层COM对象
            if (_picture != null)
                Marshal.ReleaseComObject(_picture);
            _picture = null;
            _interior = null;
            _excelShape = null;
            _border = null;
            _topLeftCellRange = null;
            _bottomRightCellRange = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置图片的名称
    /// </summary>
    public string Name
    {
        get => _picture?.Name?.ToString();
        set
        {
            if (_picture != null && value != null)
                _picture.Name = value;
        }
    }

    /// <summary>
    /// 获取图片的索引位置
    /// </summary>
    public int Index => _picture?.Index ?? 0;

    public bool Visible
    {
        get => _picture != null && _picture.Visible;
        set
        {
            if (_picture != null)
                _picture.Visible = value;
        }
    }

    public bool PrintObject
    {
        get => _picture != null && _picture.PrintObject;
        set
        {
            if (_picture != null)
                _picture.PrintObject = value;
        }
    }

    public bool Locked
    {
        get => _picture != null && _picture.Locked;
        set
        {
            if (_picture != null)
                _picture.Locked = value;
        }
    }

    public bool Enabled
    {
        get => _picture != null && _picture.Enabled;
        set
        {
            if (_picture != null)
                _picture.Enabled = value;
        }
    }

    /// <summary>
    /// 获取图片所在的父对象
    /// </summary>
    public object Parent => _picture?.Parent;

    private IExcelRange _topLeftCellRange;

    public IExcelRange TopLeftCell
    {
        get
        {
            _topLeftCellRange ??= new ExcelRange(_picture.TopLeftCell);
            return _topLeftCellRange;
        }
    }

    private IExcelRange _bottomRightCellRange;

    public IExcelRange BottomRightCell
    {
        get
        {
            _bottomRightCellRange ??= new ExcelRange(_picture.BottomRightCell);
            return _bottomRightCellRange;
        }
    }

    private IExcelInterior _interior;

    public IExcelInterior Interior
    {
        get
        {
            _interior ??= new ExcelInterior(_picture.Interior);
            return _interior;
        }
    }

    private IExcelBorder _border;

    public IExcelBorder Border
    {
        get
        {
            _border ??= new ExcelBorder(_picture.Border);
            return _border;
        }
    }

    /// <summary>
    /// 形状对象缓存
    /// </summary>
    private IExcelShapeRange _excelShape;

    /// <summary>
    /// 获取图片的底层形状对象
    /// </summary>
    public IExcelShapeRange ShapeRange
    {
        get
        {
            _excelShape ??= new ExcelShapeRange(_picture.ShapeRange);
            return _excelShape;
        }
    }

    public int ZOrder
    {
        get => _picture.ZOrder;
    }
    #endregion

    #region 位置和大小

    /// <summary>
    /// 获取或设置图片的左边距
    /// </summary>
    public double Left
    {
        get => _picture?.Left ?? 0;
        set
        {
            if (_picture != null)
                _picture.Left = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置图片的顶边距
    /// </summary>
    public double Top
    {
        get => _picture?.Top ?? 0;
        set
        {
            if (_picture != null)
                _picture.Top = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置图片的宽度
    /// </summary>
    public double Width
    {
        get => _picture?.Width ?? 0;
        set
        {
            if (_picture != null)
                _picture.Width = (float)value;
        }
    }

    /// <summary>
    /// 获取或设置图片的高度
    /// </summary>
    public double Height
    {
        get => _picture?.Height ?? 0;
        set
        {
            if (_picture != null)
                _picture.Height = (float)value;
        }
    }

    #endregion

    #region 图片属性



    /// <summary>
    /// 获取图片的原始宽度
    /// </summary>
    public double OriginalWidth => _picture?.Width ?? 0;

    /// <summary>
    /// 获取图片的原始高度
    /// </summary>
    public double OriginalHeight => _picture?.Height ?? 0;

    /// <summary>
    /// 获取图片的纵横比
    /// </summary>
    public double AspectRatio
    {
        get
        {
            double width = Width;
            double height = Height;
            return height != 0 ? width / height : 1;
        }
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 选择图片
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        _picture?.Select(replace);
    }


    /// <summary>
    /// 删除图片
    /// </summary>
    public void Delete()
    {
        _picture?.Delete();
    }

    public IExcelPicture? Copy()
    {
        var obj = _picture?.Copy();
        if (obj != null && obj is MsExcel.Picture pic)
            return new ExcelPicture(pic);
        return null;
    }

    public IExcelPicture? CopyPicture(
        XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter,
        XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture)
    {
        var obj = _picture?.CopyPicture((MsExcel.XlPictureAppearance)(int)Appearance, (MsExcel.XlCopyPictureFormat)(int)Format);
        if (obj != null && obj is MsExcel.Picture pic)
            return new ExcelPicture(pic);
        return null;
    }

    public IExcelPicture? Cut()
    {
        var obj = _picture?.Cut();
        if (obj != null && obj is MsExcel.Picture pic)
            return new ExcelPicture(pic);
        return null;
    }

    /// <summary>
    /// 调整图片大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    /// <param name="keepAspectRatio">是否保持纵横比</param>
    public void Resize(double width, double height, bool keepAspectRatio = true)
    {
        if (_picture == null) return;

        try
        {
            if (keepAspectRatio)
            {
                double originalRatio = AspectRatio;
                double newRatio = width / height;

                if (newRatio > originalRatio)
                {
                    // 以高度为准
                    width = height * originalRatio;
                }
                else
                {
                    // 以宽度为准
                    height = width / originalRatio;
                }
            }

            _picture.Width = (float)width;
            _picture.Height = (float)height;
        }
        catch
        {
            // 忽略调整大小过程中的异常
        }
    }

    /// <summary>
    /// 移动图片
    /// </summary>
    /// <param name="left">新左边距</param>
    /// <param name="top">新顶边距</param>
    public void Move(double left, double top)
    {
        if (_picture == null) return;

        try
        {
            _picture.Left = (float)left;
            _picture.Top = (float)top;
        }
        catch
        {
            // 忽略移动过程中的异常
        }
    }



    #endregion

    #region 图像处理 
    /// <summary>
    /// 按比例缩放图片
    /// </summary>
    /// <param name="scale">缩放比例</param>
    public void Scale(double scale)
    {
        if (_picture == null || scale <= 0) return;

        try
        {
            double newWidth = OriginalWidth * scale;
            double newHeight = OriginalHeight * scale;
            Resize(newWidth, newHeight, true);
        }
        catch
        {
            // 忽略缩放过程中的异常
        }
    }
    #endregion    
}