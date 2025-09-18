//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Line (边框线条) 对象的二次封装实现类
/// 实现 IExcelLine 接口
/// </summary>
internal class ExcelLine : IExcelLine
{
    #region 私有字段

    /// <summary>
    /// 内部持有的 Microsoft.Office.Interop.Excel.Line 对象引用
    /// </summary>
    private MsExcel.Line _line;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 ExcelLine 实例
    /// </summary>
    /// <param name="line">要封装的 Microsoft.Office.Interop.Excel.Line 对象</param>
    internal ExcelLine(MsExcel.Line line)
    {
        _line = line ?? throw new ArgumentNullException(nameof(line));
    }

    #endregion

    #region 基础属性 (IExcelLine)

    /// <summary>
    /// 获取线条所在的父对象
    /// </summary>
    public object Parent => _line.Parent;

    public IExcelApplication Application
    {
        get
        {
            return _line?.Application is MsExcel.Application application ? new ExcelApplication(application) : null;
        }
    }

    #endregion

    #region 线条属性 (IExcelLine)
    public string Name
    {
        get
        {
            return _line.Name;
        }
        set
        {
            _line.Name = value;
        }
    }

    public bool PrintObject
    {
        get
        {
            return _line.PrintObject;
        }
        set
        {
            _line.PrintObject = value;
        }
    }

    public bool Locked
    {
        get
        {
            return _line.Locked;
        }
        set
        {
            _line.Locked = value;
        }
    }


    /// <summary>
    /// 获取或设置线条的粗细
    /// </summary>
    public double Width
    {
        get
        {
            return _line.Width;
        }
        set
        {
            _line.Width = value;
        }
    }

    public double Height
    {
        get
        {
            return _line.Height;
        }
        set
        {
            _line.Height = value;
        }
    }

    public int Index
    {
        get
        {
            return _line.Index;
        }
    }

    public int ZOrder
    {
        get
        {
            return _line.ZOrder;
        }

    }

    public double Top
    {
        get
        {
            return _line.Top;
        }
        set
        {
            _line.Top = value;
        }
    }

    public double Left
    {
        get
        {
            return _line.Left;
        }
        set
        {
            _line.Left = value;
        }
    }

    public XlArrowHeadStyle ArrowHeadStyle
    {
        get => _line != null ? (XlArrowHeadStyle)Enum.ToObject(typeof(XlArrowHeadStyle), _line.ArrowHeadStyle) : XlArrowHeadStyle.xlArrowHeadStyleClosed;
        set
        {
            if (_line != null)
                _line.ArrowHeadStyle = (MsExcel.XlArrowHeadStyle)Enum.ToObject(typeof(MsExcel.XlArrowHeadStyle), (int)value);
        }
    }

    public XlArrowHeadWidth ArrowHeadWidth
    {
        get
        {
            return _line != null ? (XlArrowHeadWidth)Enum.ToObject(typeof(XlArrowHeadWidth), _line.ArrowHeadWidth) : XlArrowHeadWidth.xlArrowHeadWidthMedium;
        }
        set
        {
            if (_line != null)
                _line.ArrowHeadWidth = (MsExcel.XlArrowHeadWidth)Enum.ToObject(typeof(MsExcel.XlArrowHeadWidth), (int)value);
        }
    }

    public float ArrowHeadLength
    {
        get
        {
            return _line.ArrowHeadLength.ConvertToFloat();
        }
        set
        {
            _line.ArrowHeadLength = value;
        }
    }

    public IExcelRange? TopLeftCell
    {
        get
        {
            return _line.TopLeftCell is MsExcel.Range range ? new ExcelRange(range) : null;
        }
    }

    public IExcelRange? BottomRightCell
    {
        get
        {
            return _line.BottomRightCell is MsExcel.Range range ? new ExcelRange(range) : null;
        }
    }

    public IExcelShapeRange? ShapeRange
    {
        get
        {
            return _line.ShapeRange is MsExcel.ShapeRange shapeRange ? new ExcelShapeRange(shapeRange) : null;
        }
    }

    public IExcelBorder? Border
    {
        get
        {
            return _line.Border is MsExcel.Border border ? new ExcelBorder(border) : null;
        }
    }

    /// <summary>
    /// 获取或设置线条是否可见
    /// </summary>
    public bool Visible
    {
        get
        {
            return _line.Visible;
        }
        set
        {
            _line.Visible = value;
        }
    }

    public bool Enabled
    {
        get
        {
            return _line.Enabled;
        }
        set
        {
            _line.Enabled = value;
        }
    }
    #endregion

    public object BringToFront()
    {
        return _line.BringToFront();
    }
    public object SendToBack()
    {
        return _line.SendToBack();
    }

    public object Cut()
    {
        return _line.Cut();
    }

    public object Copy()
    {
        return _line.Copy();
    }

    public object Delete()
    {
        return _line.Delete();
    }

    public object Duplicate()
    {
        return _line.Duplicate();
    }

    public object CopyPicture(XlPictureAppearance appearance, XlCopyPictureFormat format)
    {
        return _line.CopyPicture(appearance.EnumConvert(MsExcel.XlPictureAppearance.xlScreen), format.EnumConvert(MsExcel.XlCopyPictureFormat.xlPicture));
    }

    public object Select(bool replace = true)
    {
        return _line.Select(replace);
    }

    #region IDisposable Support
    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_line != null)
                Marshal.ReleaseComObject(_line);
            _line = null;
        }
        _disposedValue = true;
    }

    /// <summary>
    /// 终结器 (析构函数)
    /// </summary>
    ~ExcelLine()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，实现 IDisposable 接口
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}