//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordTextColumns"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.TextColumns 对象。
/// </summary>
internal class WordTextColumns : IWordTextColumns
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTextColumns));
    internal MsWord.TextColumns _textColumns;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordTextColumns"/> 类的新实例。
    /// </summary>
    /// <param name="textColumns">原始的 Microsoft.Office.Interop.Word.TextColumns 对象。</param>
    /// <exception cref="ArgumentNullException">如果 <paramref name="textColumns"/> 为 null。</exception>
    internal WordTextColumns(MsWord.TextColumns textColumns)
    {
        _textColumns = textColumns ?? throw new ArgumentNullException(nameof(textColumns));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _textColumns?.Application != null ? new WordApplication(_textColumns.Application) : null;

    /// <inheritdoc/>
    public int Creator => _textColumns?.Creator ?? 0;

    /// <inheritdoc/>
    public object Parent => _textColumns?.Parent;

    /// <inheritdoc/>
    public int Count => _textColumns?.Count ?? 0;

    /// <inheritdoc/>
    public float? Spacing
    {
        get => _textColumns?.Spacing;
        set { if (_textColumns != null && value.HasValue) _textColumns.Spacing = value.Value; }
    }

    /// <inheritdoc/>
    public int? LineBetween
    {
        get => _textColumns?.LineBetween;
        set { if (_textColumns != null && value.HasValue) _textColumns.LineBetween = value.Value; }
    }

    /// <inheritdoc/>
    public IWordTextColumn this[int index]
    {
        get
        {
            if (_disposedValue || _textColumns == null)
            {
                return null;
            }
            if (index < 1 || index > Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");
            }
            try
            {
                var comTextColumn = _textColumns[index];
                return new WordTextColumn(comTextColumn);
            }
            catch (COMException ex)
            {
                log.Error($"Failed to retrieve text column at index {index}: {ex.Message}", ex);
                return null; // 或抛出自定义异常
            }
        }
    }

    #endregion // 属性实现

    #region 方法实现

    /// <inheritdoc/>
    public IWordTextColumn Add(int? width, int? spacing, int? evenlySpaced)
    {
        if (_textColumns == null) return null;

        try
        {
            var newColumn = _textColumns.Add(width.ComArgsVal(), spacing.ComArgsVal(), evenlySpaced.ComArgsVal());
            return newColumn != null ? new WordTextColumn(newColumn) : null;
        }
        catch (COMException ex)
        {
            // 处理 COM 异常，例如添加失败
            System.Diagnostics.Debug.WriteLine($"Failed to add text column: {ex.Message}");
            return null; // 或抛出自定义异常
        }
    }

    /// <inheritdoc/>
    public void SetCount(int count)
    {
        if (_textColumns == null) return;

        try
        {
            _textColumns.SetCount(count);
        }
        catch (COMException ex)
        {
            log.Error($"Failed to set text column count to {count}: {ex.Message}", ex);
            throw; // Re-throw or handle as appropriate
        }
    }

    #endregion // 方法实现

    #region IEnumerable<IWordTextColumn> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordTextColumn> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            var column = this[i];
            if (column != null)
            {
                yield return column;
            }
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion 

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordTextColumns"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_textColumns != null)
                {
                    Marshal.ReleaseComObject(_textColumns);
                    _textColumns = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordTextColumns"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}