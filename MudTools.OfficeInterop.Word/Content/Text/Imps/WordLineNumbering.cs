//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordLineNumbering"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.LineNumbering 对象。
/// </summary>
internal class WordLineNumbering : IWordLineNumbering
{
    internal MsWord.LineNumbering _lineNumbering;
    internal MsWord.LineNumbering InternalComObject => _lineNumbering;

    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordLineNumbering"/> 类的新实例。
    /// </summary>
    /// <param name="lineNumbering">原始的 Microsoft.Office.Interop.Word.LineNumbering 对象。</param>
    /// <exception cref="ArgumentNullException">如果 <paramref name="lineNumbering"/> 为 null。</exception>
    internal WordLineNumbering(MsWord.LineNumbering lineNumbering)
    {
        _lineNumbering = lineNumbering ?? throw new ArgumentNullException(nameof(lineNumbering));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _lineNumbering?.Application != null ? new WordApplication(_lineNumbering.Application) : null;

    /// <inheritdoc/>
    public int Creator => _lineNumbering?.Creator ?? 0;

    /// <inheritdoc/>
    public object? Parent => _lineNumbering?.Parent;

    /// <inheritdoc/>
    public int? StartingNumber
    {
        get => _lineNumbering?.StartingNumber;
        set { if (_lineNumbering != null && value.HasValue) _lineNumbering.StartingNumber = value.Value; }
    }

    /// <inheritdoc/>
    public int? CountBy
    {
        get => _lineNumbering?.CountBy;
        set { if (_lineNumbering != null && value.HasValue) _lineNumbering.CountBy = value.Value; }
    }

    /// <inheritdoc/>
    public WdNumberingRule RestartMode
    {
        get => _lineNumbering != null ? (WdNumberingRule)_lineNumbering.RestartMode : WdNumberingRule.wdRestartContinuous; // 默认值
        set
        {
            if (_lineNumbering != null) _lineNumbering.RestartMode = (MsWord.WdNumberingRule)(int)value;
        }
    }

    /// <inheritdoc/>
    public float? DistanceFromText
    {
        get => _lineNumbering?.DistanceFromText;
        set { if (_lineNumbering != null && value.HasValue) _lineNumbering.DistanceFromText = value.Value; }
    }


    /// <inheritdoc/>
    /// <remarks>
    /// 注意：此属性可能与 RestartMode 属性相关联或重叠。
    /// 具体行为取决于 Word 版本和上下文。建议优先使用 RestartMode。
    /// </remarks>
    public int? Active
    {
        get => _lineNumbering?.Active;
        set { if (_lineNumbering != null && value.HasValue) _lineNumbering.Active = value.Value; }
    }

    #endregion // 属性实现

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordLineNumbering"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_lineNumbering != null)
                {
                    Marshal.ReleaseComObject(_lineNumbering);
                    _lineNumbering = null;
                }
            }

            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordLineNumbering"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}