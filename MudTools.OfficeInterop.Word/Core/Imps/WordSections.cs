//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// 表示文档中所有节集合的封装实现类。
/// </summary>
internal class WordSections : IWordSections
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordSections));
    private MsWord.Sections _sections;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordSections"/> 类的新实例。
    /// </summary>
    /// <param name="sections">要封装的原始 COM Sections 对象。</param>
    internal WordSections(MsWord.Sections sections)
    {
        _sections = sections ?? throw new ArgumentNullException(nameof(sections));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _sections != null ? new WordApplication(_sections.Application) : null;

    /// <inheritdoc/>
    public object Parent => _sections?.Parent;

    /// <inheritdoc/>
    public int Creator => _sections?.Creator ?? 0;

    /// <inheritdoc/>
    public int Count => _sections?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordSection this[int index]
    {
        get
        {
            if (_sections == null || index < 1 || index > Count) return null;
            try
            {
                var comSection = _sections[index];
                return comSection != null ? new WordSection(comSection) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 节集合方法实现 (Sections Collection Methods Implementation)

    /// <inheritdoc/>
    public IWordSection? Add(IWordRange range, WdSectionStart start)
    {
        if (_sections == null) return null;
        try
        {
            var newSection = _sections.Add(((WordRange)range)._range, (MsWord.WdSectionStart)(int)start);
            return newSection != null ? new WordSection(newSection) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add section: {ex.Message}", ex);
            return null;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordSections"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _sections != null)
        {
            Marshal.ReleaseComObject(_sections);
            _sections = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordSections"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordSection> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordSection> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}

