//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel Axis 对象的二次封装实现类
/// 实现 IExcelAxis 接口
/// </summary>
internal class ExcelAxis : IExcelAxis
{
    #region 私有字段

    /// <summary>
    /// 内部持有的 Microsoft.Office.Interop.Excel.Axis 对象引用
    /// </summary>
    private MsExcel.Axis? _axis;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 ExcelAxis 实例
    /// </summary>
    /// <param name="axis">要封装的 Microsoft.Office.Interop.Excel.Axis 对象</param>
    /// <exception cref="ArgumentNullException">当 axis 为 null 时抛出</exception>
    internal ExcelAxis(MsExcel.Axis axis)
    {
        _axis = axis ?? throw new ArgumentNullException(nameof(axis));
    }

    #endregion

    #region 基础属性 (IExcelAxis)   

    /// <summary>
    /// 获取坐标轴的父对象
    /// </summary>
    public object Parent => _axis.Parent;

    /// <summary>
    /// 获取坐标轴所在的 Application 对象
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_axis.Application);

    #endregion

    #region 坐标轴属性 (IExcelAxis)

    /// <summary>
    /// 获取坐标轴的类型
    /// </summary>
    public int AxisType => (int)_axis.Type;

    /// <summary>
    /// 获取坐标轴的分组
    /// </summary>
    public int AxisGroup => (int)_axis.AxisGroup;

    /// <summary>
    /// 获取或设置坐标轴标题
    /// </summary>
    public string AxisTitle
    {
        get
        {
            if (_axis.HasTitle)
            {
                return _axis.AxisTitle.Text;
            }
            return null;
        }
        set
        {
            if (!string.IsNullOrEmpty(value))
            {
                _axis.HasTitle = true;
                _axis.AxisTitle.Text = value;
            }
            else
            {
                _axis.HasTitle = false;
            }
        }
    }


    /// <summary>
    /// 获取或设置坐标轴的位置类型
    /// </summary>
    public int Crosses
    {
        get => (int)_axis.Crosses;
        set => _axis.Crosses = (MsExcel.XlAxisCrosses)value;
    }

    /// <summary>
    /// 获取或设置坐标轴在指定数值处穿过另一轴
    /// </summary>
    public double CrossesAt
    {
        get => (double)_axis.CrossesAt;
        set => _axis.CrossesAt = value;
    }

    /// <summary>
    /// 获取或设置坐标轴是否在刻度线之间（类别轴）
    /// </summary>
    public bool AxisBetweenCategories
    {
        get => _axis.AxisBetweenCategories;
        set => _axis.AxisBetweenCategories = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的最小值
    /// </summary>
    public double MinimumScale
    {
        get => (double)_axis.MinimumScale;
        set => _axis.MinimumScale = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的最大值
    /// </summary>
    public double MaximumScale
    {
        get => (double)_axis.MaximumScale;
        set => _axis.MaximumScale = value;
    }

    /// <summary>
    /// 获取或设置坐标轴主要刻度单位
    /// </summary>
    public double MajorUnit
    {
        get => (double)_axis.MajorUnit;
        set => _axis.MajorUnit = value;
    }

    /// <summary>
    /// 获取或设置坐标轴次要刻度单位
    /// </summary>
    public double MinorUnit
    {
        get => (double)_axis.MinorUnit;
        set => _axis.MinorUnit = value;
    }

    /// <summary>
    /// 获取或设置坐标轴主要刻度线的类型
    /// </summary>
    public XlTickMark MajorTickMark
    {
        get => (XlTickMark)_axis.MajorTickMark;
        set => _axis.MajorTickMark = (MsExcel.XlTickMark)value;
    }

    /// <summary>
    /// 获取或设置坐标轴次要刻度线的类型
    /// </summary>
    public XlTickMark MinorTickMark
    {
        get => (XlTickMark)_axis.MinorTickMark;
        set => _axis.MinorTickMark = (MsExcel.XlTickMark)value;
    }

    /// <summary>
    /// 获取或设置坐标轴标签的位置
    /// </summary>
    public XlTickLabelPosition TickLabelPosition
    {
        get => (XlTickLabelPosition)_axis.TickLabelPosition;
        set => _axis.TickLabelPosition = (MsExcel.XlTickLabelPosition)value;
    }

    /// <summary>
    /// 获取或设置坐标轴标签的方向（角度）
    /// </summary>
    public XlTickLabelOrientation TickLabelOrientation
    {
        get => (XlTickLabelOrientation)_axis.TickLabels.Orientation;
        set => _axis.TickLabels.Orientation = (MsExcel.XlTickLabelOrientation)value;

    }

    /// <summary>
    /// 获取或设置坐标轴标签的数字格式
    /// </summary>
    public string TickLabelNumberFormat
    {
        get
        {
            try
            {
                return _axis.TickLabels?.NumberFormat ?? "";
            }
            catch { return ""; }
        }
        set
        {
            try
            {
                var tickLabels = _axis.TickLabels;
                if (tickLabels != null)
                {
                    tickLabels.NumberFormat = value;
                }
            }
            catch
            {

            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴是否反转刻度值
    /// </summary>
    public bool ReversePlotOrder
    {
        get => _axis.ReversePlotOrder;
        set => _axis.ReversePlotOrder = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的对数刻度底数
    /// </summary>
    public double LogBase
    {
        get => _axis.LogBase;
        set => _axis.LogBase = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的主要单位是否自动确定
    /// </summary>
    public bool MajorUnitIsAuto
    {
        get => _axis.MajorUnitIsAuto;
        set => _axis.MajorUnitIsAuto = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的次要单位是否自动确定
    /// </summary>
    public bool MinorUnitIsAuto
    {
        get => _axis.MinorUnitIsAuto;
        set => _axis.MinorUnitIsAuto = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的最小刻度值是否自动确定
    /// </summary>
    public bool MinimumScaleIsAuto
    {
        get => _axis.MinimumScaleIsAuto;
        set => _axis.MinimumScaleIsAuto = value;
    }

    /// <summary>
    /// 获取或设置坐标轴的最大刻度值是否自动确定
    /// </summary>
    public bool MaximumScaleIsAuto
    {
        get => _axis.MaximumScaleIsAuto;
        set => _axis.MaximumScaleIsAuto = value;
    }
    #endregion

    #region 格式设置 (IExcelAxis)

    /// <summary>
    /// 获取坐标轴标题的字体对象
    /// </summary>
    public IExcelFont? TitleFont
    {
        get
        {
            if (_axis.HasTitle)
            {
                return new ExcelFont(_axis.AxisTitle.Font);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取坐标轴刻度线标签的字体对象
    /// </summary>
    public IExcelFont? TickLabelFont
    {
        get
        {
            var tickLabels = _axis.TickLabels;
            if (tickLabels != null)
            {
                return new ExcelFont(tickLabels.Font);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取坐标轴刻度线标签对象
    /// </summary>
    public IExcelTickLabels TickLabels => new ExcelTickLabels(_axis.TickLabels); // 假设 ExcelTickLabels 存在

    /// <summary>
    /// 获取坐标轴的主要网格线对象
    /// </summary>
    public IExcelGridlines MajorGridlines
    {
        get
        {
            var gridlines = _axis.MajorGridlines;
            if (gridlines != null)
            {
                return new ExcelGridlines(gridlines); // 假设 ExcelGridlines 存在
            }
            return null;
        }
    }

    /// <summary>
    /// 获取坐标轴的次要网格线对象
    /// </summary>
    public IExcelGridlines MinorGridlines
    {
        get
        {
            var gridlines = _axis.MinorGridlines;
            if (gridlines != null)
            {
                return new ExcelGridlines(gridlines); // 假设 ExcelGridlines 存在
            }
            return null;
        }
    }
    #endregion

    #region 操作方法 (IExcelAxis)    

    /// <summary>
    /// 删除坐标轴（通常不直接删除）
    /// </summary>
    public void Delete()
    {
        try
        {
            _axis.Delete();
        }
        catch
        {
        }
    }
    #endregion

    #region IDisposable Support

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_axis != null)
                    Marshal.ReleaseComObject(_axis);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _axis = null;
        }
        _disposedValue = true;

    }

    /// <summary>
    /// 终结器 (析构函数)，防止资源未被释放
    /// </summary>
    ~ExcelAxis()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，实现 IDisposable 接口
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        // 抑制终结器
        GC.SuppressFinalize(this);
    }
    #endregion
}
