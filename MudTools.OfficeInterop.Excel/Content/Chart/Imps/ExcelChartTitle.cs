//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ChartTitle 对象的二次封装实现类
/// 实现 IExcelChartTitle 接口
/// </summary>
internal class ExcelChartTitle : IExcelChartTitle
{
    private MsExcel.ChartTitle? _chartTitle;
    private bool _disposedValue = false;

    internal ExcelChartTitle(MsExcel.ChartTitle chartTitle)
    {
        _chartTitle = chartTitle ?? throw new ArgumentNullException(nameof(chartTitle));
    }

    #region 基础属性
    public string Name
    {
        get => _chartTitle != null ? _chartTitle.Name : string.Empty;
    }


    public string Text
    {
        get => _chartTitle != null ? _chartTitle.Text : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Text = value;
        }
    }

    public string Caption
    {
        get => _chartTitle != null ? _chartTitle.Caption : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Caption = value;
        }
    }

    public object? Parent => _chartTitle?.Parent;

    public IExcelApplication? Application => _chartTitle != null ? new ExcelApplication(_chartTitle.Application) : null;
    #endregion

    #region 位置和大小
    public double Left
    {
        get => _chartTitle != null ? _chartTitle.Left : 0;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Left = value;
        }
    }

    public double Top
    {
        get => _chartTitle != null ? _chartTitle.Top : 0;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Top = value;
        }
    }

    public double Width
    {
        get => _chartTitle != null ? _chartTitle.Width : 0;
    }

    public double Height
    {
        get => _chartTitle != null ? _chartTitle.Height : 0;
    }
    #endregion

    #region 格式设置
    public bool AutoScaleFont
    {
        get => _chartTitle != null ? _chartTitle.AutoScaleFont.ConvertToBool() : false;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.AutoScaleFont = value;
        }
    }

    public bool IncludeInLayout
    {
        get => _chartTitle != null ? _chartTitle.IncludeInLayout : false;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.IncludeInLayout = value;
        }
    }

    public bool Shadow
    {
        get => _chartTitle != null ? _chartTitle.Shadow : false;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Shadow = value;
        }
    }

    public string Formula
    {
        get => _chartTitle != null ? _chartTitle.Formula : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Formula = value;
        }
    }

    public string FormulaR1C1
    {
        get => _chartTitle != null ? _chartTitle.FormulaR1C1 : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.FormulaR1C1 = value;
        }
    }

    public string FormulaLocal
    {
        get => _chartTitle != null ? _chartTitle.FormulaLocal : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.FormulaLocal = value;
        }
    }

    public string FormulaR1C1Local
    {
        get => _chartTitle != null ? _chartTitle.FormulaR1C1Local : string.Empty;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.FormulaR1C1Local = value;
        }
    }

    public XlChartElementPosition Position
    {
        get => _chartTitle != null ? _chartTitle.Position.ObjectConvertEnum(XlChartElementPosition.xlChartElementPositionAutomatic) : XlChartElementPosition.xlChartElementPositionAutomatic;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Position = value.EnumConvert(MsExcel.XlChartElementPosition.xlChartElementPositionAutomatic);
        }
    }

    public IExcelFont? Font
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelFont(_chartTitle.Font);
        }
    }

    public IExcelChartFormat? Format
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelChartFormat(_chartTitle.Format);
        }
    }

    public IExcelBorder? Border
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelBorder(_chartTitle.Border);
        }
    }

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior? Interior
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelInterior(_chartTitle.Interior);
        }
    }

    public IExcelChartFillFormat? Fill
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelChartFillFormat(_chartTitle.Fill);
        }
    }

    public IExcelCharacters? Characters
    {
        get
        {
            if (_chartTitle == null)
                return null;
            return new ExcelCharacters(_chartTitle.Characters);
        }

    }


    public int ReadingOrder
    {
        get => _chartTitle != null ? _chartTitle.ReadingOrder : 0;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.ReadingOrder = value;
        }
    }

    public XlOrientation Orientation
    {
        get => _chartTitle != null ? _chartTitle.Orientation.ObjectConvertEnum(XlOrientation.xlUpward) : XlOrientation.xlUpward;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.Orientation = value.EnumConvert(MsExcel.XlOrientation.xlUpward);
        }
    }

    public XlHAlign HorizontalAlignment
    {
        get => _chartTitle != null ? _chartTitle.HorizontalAlignment.ObjectConvertEnum(XlHAlign.xlHAlignGeneral) : XlHAlign.xlHAlignGeneral;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.HorizontalAlignment = value.EnumConvert(MsExcel.XlHAlign.xlHAlignGeneral);
        }
    }

    public XlVAlign VerticalAlignment
    {
        get => _chartTitle != null ? _chartTitle.VerticalAlignment.ObjectConvertEnum(XlVAlign.xlVAlignJustify) : XlVAlign.xlVAlignJustify;
        set
        {
            if (_chartTitle == null) return;
            _chartTitle.VerticalAlignment = value.EnumConvert(MsExcel.XlVAlign.xlVAlignJustify);
        }
    }
    #endregion

    #region 操作方法
    public void Select()
    {
        _chartTitle?.Select();
    }

    public void Delete()
    {
        _chartTitle?.Delete();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放底层COM对象
            if (_chartTitle != null)
                Marshal.ReleaseComObject(_chartTitle);
            _chartTitle = null;
        }
        _disposedValue = true;
    }

    ~ExcelChartTitle()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
