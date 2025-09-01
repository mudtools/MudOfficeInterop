//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.DataTable 的封装实现类。
/// </summary>
internal class WordDataTable : IWordDataTable
{
    private MsWord.DataTable _dataTable;
    private bool _disposedValue;

    internal WordDataTable(MsWord.DataTable dataTable)
    {
        _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _dataTable != null ? new WordApplication(_dataTable.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _dataTable?.Parent;

    /// <inheritdoc/>
    public bool HasBorderHorizontal
    {
        get => _dataTable?.HasBorderHorizontal ?? false;
        set
        {
            if (_dataTable != null)
                _dataTable.HasBorderHorizontal = value;
        }
    }

    /// <inheritdoc/>
    public bool HasBorderVertical
    {
        get => _dataTable?.HasBorderVertical ?? false;
        set
        {
            if (_dataTable != null)
                _dataTable.HasBorderVertical = value;
        }
    }

    /// <inheritdoc/>
    public bool HasBorderOutline
    {
        get => _dataTable?.HasBorderOutline ?? false;
        set
        {
            if (_dataTable != null)
                _dataTable.HasBorderOutline = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowLegendKey
    {
        get => _dataTable?.ShowLegendKey ?? false;
        set
        {
            if (_dataTable != null)
                _dataTable.ShowLegendKey = value;
        }
    }

    #endregion

    #region 对象属性实现
    /// <inheritdoc/>
    public IWordChartFont? Font => _dataTable?.Font != null ? new WordChartFont(_dataTable.Font) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _dataTable?.Border != null ? new WordChartBorder(_dataTable.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _dataTable?.Format != null ? new WordChartFormat(_dataTable.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _dataTable?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _dataTable?.Delete();
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_dataTable != null)
            {
                Marshal.ReleaseComObject(_dataTable);
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