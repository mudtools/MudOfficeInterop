//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Axis 的封装实现类。
/// </summary>
internal class WordAxis : IWordAxis
{
    private MsWord.Axis _axis;
    private bool _disposedValue;

    internal WordAxis(MsWord.Axis axis)
    {
        _axis = axis ?? throw new ArgumentNullException(nameof(axis));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _axis != null ? new WordApplication(_axis.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _axis?.Parent;

    /// <inheritdoc/>
    public XlAxisType Type => _axis?.Type != null ? (XlAxisType)(int)_axis?.Type : XlAxisType.xlValue;

    /// <inheritdoc/>
    public XlAxisGroup AxisGroup => _axis?.AxisGroup != null ? (XlAxisGroup)(int)_axis?.AxisGroup : XlAxisGroup.xlPrimary;

    /// <inheritdoc/>
    public string AxisTitle
    {
        get => _axis?.AxisTitle?.Text ?? string.Empty;
        set
        {
            if (_axis?.AxisTitle != null)
                _axis.AxisTitle.Text = value;
        }
    }

    /// <inheritdoc/>
    public bool HasTitle
    {
        get => _axis?.HasTitle ?? false;
        set
        {
            if (_axis != null)
                _axis.HasTitle = value;
        }
    }

    /// <inheritdoc/>
    public bool HasMajorGridlines
    {
        get => _axis?.HasMajorGridlines ?? false;
        set
        {
            if (_axis != null)
                _axis.HasMajorGridlines = value;
        }
    }

    /// <inheritdoc/>
    public bool HasMinorGridlines
    {
        get => _axis?.HasMinorGridlines ?? false;
        set
        {
            if (_axis != null)
                _axis.HasMinorGridlines = value;
        }
    }

    /// <inheritdoc/>
    public XlTickLabelPosition TickLabelPosition
    {
        get => _axis?.TickLabelPosition != null ? (XlTickLabelPosition)(int)_axis?.TickLabelPosition : XlTickLabelPosition.xlTickLabelPositionNone;
        set
        {
            if (_axis != null) _axis.TickLabelPosition = (MsWord.XlTickLabelPosition)(int)value;
        }
    }



    /// <inheritdoc/>
    public XlTickMark MajorTickMark
    {
        get => _axis?.MajorTickMark != null ? (XlTickMark)(int)_axis?.MajorTickMark : XlTickMark.xlTickMarkNone;
        set
        {
            if (_axis != null) _axis.MajorTickMark = (MsWord.XlTickMark)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlTickMark MinorTickMark
    {
        get => _axis?.MinorTickMark != null ? (XlTickMark)(int)_axis?.MinorTickMark : XlTickMark.xlTickMarkNone;
        set
        {
            if (_axis != null) _axis.MinorTickMark = (MsWord.XlTickMark)(int)value;
        }
    }

    /// <inheritdoc/>
    public double MajorUnit
    {
        get => _axis?.MajorUnit ?? 0.0;
        set
        {
            if (_axis != null)
                _axis.MajorUnit = value;
        }
    }

    /// <inheritdoc/>
    public bool MajorUnitIsAuto
    {
        get => _axis?.MajorUnitIsAuto ?? true;
        set
        {
            if (_axis != null)
                _axis.MajorUnitIsAuto = value;
        }
    }

    /// <inheritdoc/>
    public double MinorUnit
    {
        get => _axis?.MinorUnit ?? 0.0;
        set
        {
            if (_axis != null)
                _axis.MinorUnit = value;
        }
    }

    /// <inheritdoc/>
    public bool MinorUnitIsAuto
    {
        get => _axis?.MinorUnitIsAuto ?? true;
        set
        {
            if (_axis != null)
                _axis.MinorUnitIsAuto = value;
        }
    }

    /// <inheritdoc/>
    public double MinimumScale
    {
        get => _axis?.MinimumScale ?? 0.0;
        set
        {
            if (_axis != null)
                _axis.MinimumScale = value;
        }
    }

    /// <inheritdoc/>
    public bool MinimumScaleIsAuto
    {
        get => _axis?.MinimumScaleIsAuto ?? true;
        set
        {
            if (_axis != null)
                _axis.MinimumScaleIsAuto = value;
        }
    }

    /// <inheritdoc/>
    public double MaximumScale
    {
        get => _axis?.MaximumScale ?? 0.0;
        set
        {
            if (_axis != null)
                _axis.MaximumScale = value;
        }
    }

    /// <inheritdoc/>
    public bool MaximumScaleIsAuto
    {
        get => _axis?.MaximumScaleIsAuto ?? true;
        set
        {
            if (_axis != null)
                _axis.MaximumScaleIsAuto = value;
        }
    }

    /// <inheritdoc/>
    public double CrossesAt
    {
        get => _axis?.CrossesAt ?? 0.0;
        set
        {
            if (_axis != null)
                _axis.CrossesAt = value;
        }
    }

    /// <inheritdoc/>
    public MsWord.XlAxisCrosses Crosses
    {
        get => _axis?.Crosses ?? MsWord.XlAxisCrosses.xlAxisCrossesAutomatic;
        set
        {
            if (_axis != null)
                _axis.Crosses = value;
        }
    }

    /// <inheritdoc/>
    public bool ReversePlotOrder
    {
        get => _axis?.ReversePlotOrder ?? false;
        set
        {
            if (_axis != null)
                _axis.ReversePlotOrder = value;
        }
    }

    /// <inheritdoc/>
    public bool ScaleType
    {
        get => _axis?.ScaleType == MsWord.XlScaleType.xlScaleLogarithmic;
        set
        {
            if (_axis != null)
                _axis.ScaleType = value ? MsWord.XlScaleType.xlScaleLogarithmic : MsWord.XlScaleType.xlScaleLinear;
        }
    }
    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordAxisTitle? AxisTitleObject => _axis?.AxisTitle != null ? new WordAxisTitle(_axis.AxisTitle) : null;

    /// <inheritdoc/>
    public IWordTickLabels? TickLabels => _axis?.TickLabels != null ? new WordTickLabels(_axis.TickLabels) : null;

    /// <inheritdoc/>
    public IWordGridlines? MajorGridlines => _axis?.MajorGridlines != null ? new WordGridlines(_axis.MajorGridlines) : null;

    /// <inheritdoc/>
    public IWordGridlines? MinorGridlines => _axis?.MinorGridlines != null ? new WordGridlines(_axis.MinorGridlines) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _axis?.Border != null ? new WordChartBorder(_axis.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _axis?.Format != null ? new WordChartFormat(_axis.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _axis?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _axis?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_axis != null)
            {
                Marshal.ReleaseComObject(_axis);
            }
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}