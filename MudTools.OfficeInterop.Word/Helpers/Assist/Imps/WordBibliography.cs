//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示与 Word 文档关联的书目引用功能的封装实现类。
/// </summary>
internal class WordBibliography : IWordBibliography
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordBibliography));
    private MsWord.Bibliography _bibliography;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordBibliography"/> 类的新实例。
    /// </summary>
    /// <param name="bibliography">要封装的原始 COM Bibliography 对象。</param>
    internal WordBibliography(MsWord.Bibliography bibliography)
    {
        _bibliography = bibliography ?? throw new ArgumentNullException(nameof(bibliography));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _bibliography != null ? new WordApplication(_bibliography.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _bibliography?.Parent;

    /// <inheritdoc/>
    public int Creator => _bibliography?.Creator ?? 0;

    #endregion

    #region 书目属性实现 (Bibliography Properties Implementation)

    /// <inheritdoc/>
    public IWordSources Sources => _bibliography?.Sources != null ? new WordSources(_bibliography.Sources) : null;

    #endregion

    #region 书目方法实现 (Bibliography Methods Implementation)

    /// <inheritdoc/>
    public string GenerateUniqueTag()
    {
        if (_bibliography == null) return string.Empty;
        try
        {
            return _bibliography.GenerateUniqueTag();
        }
        catch (COMException ex)
        {
            log.Error($"Failed to generate bibliography: {ex.Message}");
            return string.Empty;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordBibliography"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _bibliography != null)
        {
            Marshal.ReleaseComObject(_bibliography);
            _bibliography = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordBibliography"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}