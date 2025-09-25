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
    public bool HasTitle
    {
        get => _axis != null ? _axis.HasTitle : false;
        set
        {
            if (_axis != null)
            {
                _axis.HasTitle = value;
            }
        }
    }

    /// <summary>
    /// 获取坐标轴的类型
    /// </summary>
    public XlAxisType Type
    {
        get => _axis != null ? _axis.Type.EnumConvert(XlAxisType.xlCategory) : XlAxisType.xlCategory;
        set
        {
            if (_axis != null)
            {
                _axis.Type = value.EnumConvert(MsExcel.XlAxisType.xlCategory);
            }
        }
    }

    /// <summary>
    /// 获取坐标轴的分组
    /// </summary>
    public XlAxisGroup AxisGroup
    {
        get => _axis != null ? _axis.AxisGroup.EnumConvert(XlAxisGroup.xlPrimary) : XlAxisGroup.xlPrimary;
    }



    /// <summary>
    /// 获取或设置坐标轴的位置类型
    /// </summary>
    public XlAxisCrosses Crosses
    {
        get => _axis != null ? _axis.Crosses.EnumConvert(XlAxisCrosses.xlAxisCrossesAutomatic) : XlAxisCrosses.xlAxisCrossesAutomatic;
        set
        {
            if (_axis != null)
            {
                _axis.Crosses = value.EnumConvert(MsExcel.XlAxisCrosses.xlAxisCrossesAutomatic);
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴在指定数值处穿过另一轴
    /// </summary>
    public double CrossesAt
    {
        get => _axis != null ? _axis.CrossesAt : 0;
        set
        {
            if (_axis != null)
            {
                _axis.CrossesAt = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴是否在刻度线之间（类别轴）
    /// </summary>
    public bool AxisBetweenCategories
    {
        get => _axis != null ? _axis.AxisBetweenCategories : false;
        set
        {
            if (_axis != null)
            {
                _axis.AxisBetweenCategories = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的最小值
    /// </summary>
    public double MinimumScale
    {
        get => _axis != null ? _axis.MinimumScale : 0;
        set
        {
            if (_axis != null)
            {
                _axis.MinimumScale = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的最大值
    /// </summary>
    public double MaximumScale
    {
        get => _axis != null ? _axis.MaximumScale : 0;
        set
        {
            if (_axis != null)
            {
                _axis.MaximumScale = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴主要刻度单位
    /// </summary>
    public double MajorUnit
    {
        get => _axis != null ? _axis.MajorUnit : 0;
        set
        {
            if (_axis != null)
            {
                _axis.MajorUnit = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴次要刻度单位
    /// </summary>
    public double MinorUnit
    {
        get => _axis != null ? _axis.MinorUnit : 0;
        set
        {
            if (_axis != null)
            {
                _axis.MinorUnit = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴主要刻度线的类型
    /// </summary>
    public XlTickMark MajorTickMark
    {
        get => _axis != null ? _axis.MajorTickMark.EnumConvert(XlTickMark.xlTickMarkNone) : XlTickMark.xlTickMarkNone;
        set
        {
            if (_axis != null)
            {
                _axis.MajorTickMark = value.EnumConvert(MsExcel.XlTickMark.xlTickMarkNone);
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴次要刻度线的类型
    /// </summary>
    public XlTickMark MinorTickMark
    {
        get
        {
            if (_axis != null)
            {
                return _axis.MinorTickMark.EnumConvert(XlTickMark.xlTickMarkNone);
            }
            return XlTickMark.xlTickMarkNone;
        }
        set
        {
            if (_axis != null)
            {
                _axis.MinorTickMark = value.EnumConvert(MsExcel.XlTickMark.xlTickMarkNone);
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴标签的位置
    /// </summary>
    public XlTickLabelPosition TickLabelPosition
    {
        get
        {
            if (_axis != null)
            {
                return _axis.TickLabelPosition.EnumConvert(XlTickLabelPosition.xlTickLabelPositionHigh);
            }
            return XlTickLabelPosition.xlTickLabelPositionHigh;
        }
        set
        {
            if (_axis != null)
            {
                _axis.TickLabelPosition = value.EnumConvert(MsExcel.XlTickLabelPosition.xlTickLabelPositionHigh);
            }
        }
    }


    /// <summary>
    /// 获取或设置坐标轴是否反转刻度值
    /// </summary>
    public bool ReversePlotOrder
    {
        get
        {
            if (_axis != null)
            {
                return _axis.ReversePlotOrder;
            }
            return false;
        }
        set
        {
            if (_axis != null)
            {
                _axis.ReversePlotOrder = value;
            }
        }
    }


    /// <summary>
    /// 获取或设置坐标轴的对数刻度底数
    /// </summary>
    public double LogBase
    {
        get
        {
            if (_axis != null)
            {
                return _axis.LogBase;
            }
            return 0;
        }
        set
        {
            if (_axis != null)
            {
                _axis.LogBase = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的主要单位是否自动确定
    /// </summary>
    public bool MajorUnitIsAuto
    {
        get
        {
            if (_axis != null)
            {
                return _axis.MajorUnitIsAuto;
            }
            return false;
        }
        set
        {
            if (_axis != null)
            {
                _axis.MajorUnitIsAuto = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的次要单位是否自动确定
    /// </summary>
    public bool MinorUnitIsAuto
    {
        get
        {
            if (_axis != null)
            {
                return _axis.MinorUnitIsAuto;
            }
            return false;
        }
        set
        {
            if (_axis != null)
            {
                _axis.MinorUnitIsAuto = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的最小刻度值是否自动确定
    /// </summary>
    public bool MinimumScaleIsAuto
    {
        get
        {
            if (_axis != null)
            {
                return _axis.MinimumScaleIsAuto;
            }
            return false;
        }
        set
        {
            if (_axis != null)
            {
                _axis.MinimumScaleIsAuto = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置坐标轴的最大刻度值是否自动确定
    /// </summary>
    public bool MaximumScaleIsAuto
    {
        get
        {
            if (_axis != null)
            {
                return _axis.MaximumScaleIsAuto;
            }
            return false;
        }
        set
        {
            if (_axis != null)
            {
                _axis.MaximumScaleIsAuto = value;
            }
        }
    }
    #endregion

    #region 格式设置 (IExcelAxis)
    public IExcelAxisTitle? AxisTitle
    {
        get
        {
            if (_axis != null)
            {
                return new ExcelAxisTitle(_axis.AxisTitle);
            }
            return null;
        }
    }


    public IExcelChartFormat? Format
    {
        get
        {
            if (_axis != null)
            {
                return new ExcelChartFormat(_axis.Format);
            }
            return null;
        }
    }

    public IExcelBorder? Border
    {
        get
        {
            if (_axis != null)
            {
                return new ExcelBorder(_axis.Border);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取坐标轴刻度线标签对象
    /// </summary>
    public IExcelTickLabels? TickLabels
    {
        get
        {
            if (_axis != null)
            {
                return new ExcelTickLabels(_axis.TickLabels);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取坐标轴的主要网格线对象
    /// </summary>
    public IExcelGridlines? MajorGridlines
    {
        get
        {
            if (_axis != null)
            {
                var gridlines = _axis.MajorGridlines;
                return new ExcelGridlines(gridlines);
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
            if (_axis != null)
            {
                var gridlines = _axis.MinorGridlines;
                return new ExcelGridlines(gridlines);
            }
            return null;
        }
    }
    #endregion

    #region 操作方法 (IExcelAxis)
    /// <summary>
    /// 删除坐标轴
    /// </summary>
    public void Delete()
    {
        try
        {
            _axis?.Delete();
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
            // 释放底层COM对象
            if (_axis != null)
                Marshal.ReleaseComObject(_axis);
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
