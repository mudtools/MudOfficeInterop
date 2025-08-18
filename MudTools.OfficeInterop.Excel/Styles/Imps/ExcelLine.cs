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
    /// 内部持有的 Microsoft.Office.Interop.Excel.LineFormat 对象引用
    /// </summary>
    private MsExcel.LineFormat _lineFormat;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 ExcelLine 实例
    /// </summary>
    /// <param name="lineFormat">要封装的 Microsoft.Office.Interop.Excel.LineFormat 对象</param>
    internal ExcelLine(MsExcel.LineFormat lineFormat)
    {
        _lineFormat = lineFormat ?? throw new ArgumentNullException(nameof(lineFormat));
    }

    #endregion

    #region 基础属性 (IExcelLine)

    /// <summary>
    /// 获取线条所在的父对象
    /// </summary>
    public object Parent => _lineFormat.Parent;

    /// <summary>
    /// 获取线条对象所在的 Application 对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var parent = Parent;
            if (parent is MsExcel.Chart chart)
            {
                return new ExcelApplication(chart.Application);
            }
            return null;
        }
    }

    #endregion

    #region 线条属性 (IExcelLine)

    /// <summary>
    /// 获取或设置线条的颜色 (RGB 颜色值)
    /// </summary>
    public int Color
    {
        get
        {
            try
            {
                return _lineFormat.ForeColor.RGB;
            }
            catch { }
            return 0; // Default black or error
        }
        set
        {
            try
            {
                _lineFormat.ForeColor.RGB = value;
            }
            catch { }
        }
    }

    /// <summary>
    /// 获取或设置线条的样式
    /// </summary>
    public MsoLineStyle Style
    {
        get
        {
            try
            {
                return (MsoLineStyle)_lineFormat.Style;
            }
            catch { }
            return MsoLineStyle.msoLineSingle; // Default
        }
        set
        {
            try
            {
                _lineFormat.Style = (MsCore.MsoLineStyle)value;
            }
            catch { }
        }
    }

    /// <summary>
    /// 获取或设置线条的粗细
    /// </summary>
    public float Weight
    {
        get
        {
            try
            {
                return _lineFormat.Weight;
            }
            catch { }
            return 1.0f; // Default weight
        }
        set
        {
            try
            {
                _lineFormat.Weight = value;
            }
            catch { }
        }
    }

    /// <summary>
    /// 获取或设置线条是否可见
    /// </summary>
    public bool Visible
    {
        get
        {
            try
            {
                return _lineFormat.Visible == MsCore.MsoTriState.msoTrue;
            }
            catch { }
            return false;
        }
        set
        {
            try
            {
                _lineFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            }
            catch { }
        }
    }

    /// <summary>
    /// 获取或设置线条的透明度
    /// </summary>
    public float Transparency
    {
        get
        {
            try
            {
                return _lineFormat.Transparency;
            }
            catch { }
            return 0.0f; // Default opaque
        }
        set
        {
            try
            {
                _lineFormat.Transparency = value;
            }
            catch { }
        }
    }

    #endregion


    #region IDisposable Support
    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_lineFormat != null)
                    Marshal.ReleaseComObject(_lineFormat);
            }
            catch { }
            _lineFormat = null;
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