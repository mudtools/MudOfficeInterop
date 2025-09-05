//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartData 的封装实现类。
/// </summary>
internal class WordChartData : IWordChartData
{
    private MsWord.ChartData _chartData;
    private bool _disposedValue;

    internal WordChartData(MsWord.ChartData chartData)
    {
        _chartData = chartData ?? throw new ArgumentNullException(nameof(chartData));
        _disposedValue = false;
    }

    #region 属性实现
    /// <inheritdoc/>
    public bool IsLinked
    {
        get => _chartData?.IsLinked ?? false;
    }

    /// <inheritdoc/>
    public object Workbook => _chartData?.Workbook;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Activate()
    {
        _chartData?.Activate();
    }

    /// <inheritdoc/>
    public void BreakLink()
    {
        _chartData?.BreakLink();
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartData != null)
        {
            Marshal.ReleaseComObject(_chartData);
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