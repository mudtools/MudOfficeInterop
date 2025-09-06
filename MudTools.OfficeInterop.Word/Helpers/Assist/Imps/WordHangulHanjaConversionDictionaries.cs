//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordHangulHanjaConversionDictionaries : IWordHangulHanjaConversionDictionaries
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordHangulHanjaConversionDictionaries));
    private MsWord.HangulHanjaConversionDictionaries _dictionaries;
    private bool _disposedValue;

    internal WordHangulHanjaConversionDictionaries(MsWord.HangulHanjaConversionDictionaries dictionaries)
    {
        _dictionaries = dictionaries ?? throw new ArgumentNullException(nameof(dictionaries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _dictionaries?.Application != null ? new WordApplication(_dictionaries.Application) : null;

    /// <summary>
    /// 获取集合中自定义词典的数量。
    /// </summary>
    public int Count => _dictionaries?.Count ?? 0;

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    public int Creator => _dictionaries?.Creator ?? 0;

    /// <summary>
    /// 获取或设置活动的自定义词典。
    /// </summary>
    public IWordDictionary? ActiveCustomDictionary
    {
        get
        {
            var comDict = _dictionaries?.ActiveCustomDictionary;
            return comDict != null ? new WordDictionary(comDict) : null;
        }
        set
        {
            if (_dictionaries != null)
            {
                var dictImpl = value as WordDictionary;
                _dictionaries.ActiveCustomDictionary = dictImpl?._dictionary;
            }
        }
    }

    /// <summary>
    /// 获取内置的韩文/汉字转换词典。
    /// </summary>
    public IWordDictionary? BuiltinDictionary
    {
        get
        {
            var comDict = _dictionaries?.BuiltinDictionary;
            return comDict != null ? new WordDictionary(comDict) : null;
        }
    }

    #endregion

    #region 索引器实现

    /// <summary>
    /// 返回集合中指定的 <see cref="IWordDictionary"/> 对象。
    /// </summary>
    /// <param name="index">要返回的单个对象。可以是代表序号位置的 Number 类型的值。</param>
    /// <returns>指定索引处的 <see cref="IWordDictionary"/> 对象。</returns>
    public IWordDictionary? this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _dictionaries == null) return null;

            try
            {
                var comDictionary = _dictionaries[index];
                var wrapper = new WordDictionary(comDictionary);
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
    /// 将新的自定义词典添加到集合中。
    /// </summary>
    /// <param name="FileName">新自定义词典的完整路径和文件名。</param>
    /// <returns>返回新添加的 <see cref="IWordDictionary"/> 对象。</returns>
    public IWordDictionary Add(string FileName)
    {
        try
        {
            var comDictionary = _dictionaries?.Add(FileName);
            if (comDictionary != null)
            {
                return new WordDictionary(comDictionary);
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

    public IEnumerator<IWordDictionary> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++) // Dictionaries 索引通常从 1 开始
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
            if (_dictionaries != null)
            {
                Marshal.ReleaseComObject(_dictionaries);
                _dictionaries = null;
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