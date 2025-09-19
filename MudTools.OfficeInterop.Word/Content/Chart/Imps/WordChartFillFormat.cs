//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// Word.ChartFillFormat 的封装实现类。
/// </summary>
internal class WordChartFillFormat : IWordChartFillFormat
{
    private MsWord.ChartFillFormat _chartFillFormat;
    private bool _disposedValue;

    internal WordChartFillFormat(MsWord.ChartFillFormat chartFillFormat)
    {
        _chartFillFormat = chartFillFormat ?? throw new ArgumentNullException(nameof(chartFillFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartFillFormat != null ? new WordApplication(_chartFillFormat.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartFillFormat?.Parent;

    /// <inheritdoc/>
    public IWordChartColorFormat ForeColor => _chartFillFormat != null ? new WordChartColorFormat(_chartFillFormat.ForeColor) : null;

    /// <inheritdoc/>
    public IWordChartColorFormat BackColor => _chartFillFormat != null ? new WordChartColorFormat(_chartFillFormat.BackColor) : null;

    /// <inheritdoc/>
    public MsoPatternType Pattern
    {
        get => _chartFillFormat?.Pattern != null ? _chartFillFormat.Pattern.EnumConvert(MsoPatternType.msoPatternMixed) : MsoPatternType.msoPatternMixed;
    }

    /// <inheritdoc/>
    public MsoFillType Type
    {
        get => _chartFillFormat?.Type != null ? _chartFillFormat.Type.EnumConvert(MsoFillType.msoFillMixed) : MsoFillType.msoFillMixed;
    }

    #endregion

    #region 方法实现  

    /// <inheritdoc/>
    public void OneColorGradient(MsoGradientStyle style, int variant, float degree)
    {
        _chartFillFormat?.OneColorGradient((MsCore.MsoGradientStyle)(int)style, variant, degree);
    }

    /// <inheritdoc/>
    public void PresetTextured(MsoPresetTexture texture)
    {
        _chartFillFormat?.PresetTextured((MsCore.MsoPresetTexture)(int)texture);
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _chartFillFormat != null)
        {
            Marshal.ReleaseComObject(_chartFillFormat);
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