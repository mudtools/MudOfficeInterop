//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordOMathAutoCorrectEntries : IWordOMathAutoCorrectEntries
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordOMathAutoCorrectEntries));

    private MsWord.OMathAutoCorrectEntries _entries;
    private bool _disposedValue;

    internal WordOMathAutoCorrectEntries(MsWord.OMathAutoCorrectEntries entries)
    {
        _entries = entries ?? throw new ArgumentNullException(nameof(entries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _entries?.Application != null ? new WordApplication(_entries.Application) : null;

    /// <summary>
    /// 获取集合中数学自动更正条目的数量。
    /// </summary>
    public int Count => _entries?.Count ?? 0;

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    public int Creator => _entries?.Creator ?? 0;

    #endregion

    #region 索引器实现

    /// <summary>
    /// 返回集合中指定的 <see cref="IWordOMathAutoCorrectEntry"/> 对象。
    /// </summary>
    /// <param name="index">要返回的单个对象。可以是代表序号位置的 Number 类型的值。</param>
    /// <returns>指定索引处的 <see cref="IWordOMathAutoCorrectEntry"/> 对象。</returns>
    public IWordOMathAutoCorrectEntry? this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _entries == null) return null;

            try
            {
                var comEntry = _entries[index];
                var wrapper = new WordOMathAutoCorrectEntry(comEntry);
                return wrapper;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    /// <summary>
    /// 返回集合中指定名称的 <see cref="IWordOMathAutoCorrectEntry"/> 对象。
    /// </summary>
    /// <param name="name">条目的名称。</param>
    /// <returns>具有指定名称的 <see cref="IWordOMathAutoCorrectEntry"/> 对象。</returns>
    public IWordOMathAutoCorrectEntry? this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name) || _entries == null) return null;

            try
            {
                var comEntry = _entries[name];
                if (comEntry == null) return null;

                var wrapper = new WordOMathAutoCorrectEntry(comEntry);
                return wrapper;
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

    /// <summary>
    /// 将新的数学自动更正条目添加到集合中。
    /// </summary>
    /// <param name="Name">新条目的名称（键入的文本）。</param>
    /// <param name="Value">新条目的值（替换的数学表达式）。</param>
    /// <returns>返回新添加的 <see cref="IWordOMathAutoCorrectEntry"/> 对象。</returns>
    public IWordOMathAutoCorrectEntry Add(string Name, string Value)
    {
        try
        {
            var comEntry = _entries?.Add(Name, Value);
            if (comEntry != null)
            {
                return new WordOMathAutoCorrectEntry(comEntry);
            }
            return null;
        }
        catch (COMException ce)
        {
            log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
            return null;
        }
    }

    #endregion

    #region 方法实现 (GetEnumerator, Dispose)

    public IEnumerator<IWordOMathAutoCorrectEntry> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++) // OMathAutoCorrectEntries 索引通常从 1 开始
        {
            yield return this[i];
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放集合本身
            if (_entries != null)
            {
                Marshal.ReleaseComObject(_entries);
                _entries = null;
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