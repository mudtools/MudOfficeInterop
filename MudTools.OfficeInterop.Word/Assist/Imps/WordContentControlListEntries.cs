//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表示下拉列表或组合框内容控件中列表项集合的封装实现类。
/// </summary>
internal class WordContentControlListEntries : IWordContentControlListEntries
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordEditors));
    private MsWord.ContentControlListEntries _listEntries;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordContentControlListEntries"/> 类的新实例。
    /// </summary>
    /// <param name="listEntries">要封装的原始 COM ContentControlListEntries 对象。</param>
    internal WordContentControlListEntries(MsWord.ContentControlListEntries listEntries)
    {
        _listEntries = listEntries ?? throw new ArgumentNullException(nameof(listEntries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listEntries != null ? new WordApplication(_listEntries.Application) : null;

    /// <inheritdoc/>
    public object Parent => _listEntries?.Parent;

    /// <inheritdoc/>
    public int Count => _listEntries?.Count ?? 0;

    /// <inheritdoc/>
    public IWordContentControlListEntry this[int index]
    {
        get
        {
            if (_listEntries == null || index < 1 || index > Count) return null;
            try
            {
                var comListEntry = _listEntries[index];
                return comListEntry != null ? new WordContentControlListEntry(comListEntry) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordContentControlListEntry Add(string displayText, string value, int index = 0)
    {
        if (_listEntries == null) return null;

        try
        {
            var newListEntry = _listEntries.Add(displayText ?? string.Empty, value ?? string.Empty, index);
            return newListEntry != null ? new WordContentControlListEntry(newListEntry) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add list entry: {ex.Message}");
            return null;
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        _listEntries?.Clear();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordContentControlListEntries"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listEntries != null)
        {
            Marshal.ReleaseComObject(_listEntries);
            _listEntries = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordContentControlListEntries"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordContentControlListEntry> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordContentControlListEntry> GetEnumerator()
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