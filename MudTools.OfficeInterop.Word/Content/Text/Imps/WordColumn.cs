//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Column 的封装实现类。
/// </summary>
internal class WordColumn : IWordColumn
{
    internal MsWord.Column _column;
    internal MsWord.Column InternalComObject => _column;
    private bool _disposedValue;

    internal WordColumn(MsWord.Column column)
    {
        _column = column ?? throw new ArgumentNullException(nameof(column));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _column != null ? new WordApplication(_column.Application) : null;

    /// <inheritdoc/>
    public object Parent => _column?.Parent;

    /// <inheritdoc/>
    public int Index => _column?.Index ?? 0;

    /// <inheritdoc/>
    public IWordTable Table => _column?.Cells?.Parent as MsWord.Table != null ? new WordTable(_column.Cells.Parent as MsWord.Table) : null;

    /// <inheritdoc/>
    public float Width
    {
        get => _column?.Width ?? 0f;
        set
        {
            if (_column != null)
                _column.Width = value;
        }
    }

    /// <inheritdoc/>
    public IWordCells Cells => _column?.Cells != null ? new WordCells(_column.Cells) : null;

    /// <inheritdoc/>
    public IWordBorders Borders => _column?.Borders != null ? new WordBorders(_column.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading => _column?.Shading != null ? new WordShading(_column.Shading) : null;


    /// <inheritdoc/>
    public float PreferredWidth
    {
        get => _column?.PreferredWidth ?? 0f;
        set
        {
            if (_column != null)
                _column.PreferredWidth = value;
        }
    }

    /// <inheritdoc/>
    public WdPreferredWidthType PreferredWidthType
    {
        get => _column?.PreferredWidthType != null ? (WdPreferredWidthType)(int)_column?.PreferredWidthType : WdPreferredWidthType.wdPreferredWidthAuto;
        set
        {
            if (_column != null) _column.PreferredWidthType = (MsWord.WdPreferredWidthType)(int)value;
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _column?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _column?.Delete();
    }


    /// <inheritdoc/>
    public void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_column?.Borders != null)
        {
            try
            {
                _column.Borders.Enable = 1;
                foreach (MsWord.Border border in _column.Borders)
                {
                    border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                    border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                    border.Color = (MsWord.WdColor)(int)color;
                }
            }
            catch
            {
                // 设置边框失败忽略异常
            }
        }
    }

    /// <inheritdoc/>
    public void RemoveBorders()
    {
        if (_column == null)
            return;
        _column.Borders.Enable = 0;
    }

    /// <inheritdoc/>
    public void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_column?.Shading != null)
        {
            try
            {
                _column.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                if (foregroundColor != WdColor.wdColorAutomatic)
                    _column.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                if (backgroundColor != WdColor.wdColorWhite)
                    _column.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;
            }
            catch
            {
                // 设置底纹失败忽略异常
            }
        }
    }

    /// <inheritdoc/>
    public void RemoveShading()
    {
        if (_column == null)
            return;
        _column.Shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_column != null)
            {
                Marshal.ReleaseComObject(_column);
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