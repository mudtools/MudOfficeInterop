//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartBorder 的封装实现类。
/// </summary>
internal class WordChartBorder : IWordChartBorder
{
    private MsWord.ChartBorder _chartBorder;
    private bool _disposedValue;

    internal WordChartBorder(MsWord.ChartBorder chartBorder)
    {
        _chartBorder = chartBorder ?? throw new ArgumentNullException(nameof(chartBorder));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartBorder != null ? new WordApplication(_chartBorder.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartBorder?.Parent;

    /// <inheritdoc/>
    public object Color
    {
        get => _chartBorder?.Color;
        set
        {
            if (_chartBorder != null)
                _chartBorder.Color = value;
        }
    }

    /// <inheritdoc/>
    public XlColorIndex ColorIndex
    {
        get => _chartBorder?.ColorIndex != null ? (XlColorIndex)(int)_chartBorder?.ColorIndex : XlColorIndex.xlColorIndexNone;
        set
        {
            if (_chartBorder != null) _chartBorder.ColorIndex = (MsCore.XlColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlLineStyle LineStyle
    {
        get => _chartBorder?.LineStyle != null ? (XlLineStyle)(int)_chartBorder?.LineStyle : XlLineStyle.xlLineStyleNone;
        set
        {
            if (_chartBorder != null) _chartBorder.LineStyle = (MsWord.XlLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public float Weight
    {
        get => _chartBorder?.Weight.ConvertToFloat() ?? 0f;
        set
        {
            if (_chartBorder != null)
                _chartBorder.Weight = value;
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartBorder != null)
        {
            Marshal.ReleaseComObject(_chartBorder);
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