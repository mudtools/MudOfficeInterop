//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordOMathRecognizedFunctions : IWordOMathRecognizedFunctions
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordOMathRecognizedFunctions));

    private MsWord.OMathRecognizedFunctions _functions;
    private bool _disposedValue;

    internal WordOMathRecognizedFunctions(MsWord.OMathRecognizedFunctions functions)
    {
        _functions = functions ?? throw new ArgumentNullException(nameof(functions));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _functions?.Application != null ? new WordApplication(_functions.Application) : null;

    /// <summary>
    /// 获取集合中数学识别函数的数量。
    /// </summary>
    public int Count => _functions?.Count ?? 0;

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    public int Creator => _functions?.Creator ?? 0;

    #endregion

    #region 索引器实现

    /// <summary>
    /// 返回集合中指定的 <see cref="IWordOMathRecognizedFunction"/> 对象。
    /// </summary>
    /// <param name="index">要返回的单个对象。可以是代表序号位置的 Number 类型的值。</param>
    /// <returns>指定索引处的 <see cref="IWordOMathRecognizedFunction"/> 对象。</returns>
    public IWordOMathRecognizedFunction? this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _functions == null) return null;

            try
            {
                var comFunction = _functions[index];
                var wrapper = new WordOMathRecognizedFunction(comFunction);
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
    /// 将新的数学识别函数添加到集合中。
    /// </summary>
    /// <param name="Name">要添加的函数名称。</param>
    /// <returns>返回新添加的 <see cref="IWordOMathRecognizedFunction"/> 对象。</returns>
    public IWordOMathRecognizedFunction Add(string Name)
    {
        try
        {
            var comFunction = _functions?.Add(Name);
            if (comFunction != null)
            {
                return new WordOMathRecognizedFunction(comFunction);
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

    public IEnumerator<IWordOMathRecognizedFunction> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++) // OMathRecognizedFunctions 索引通常从 1 开始
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
            if (_functions != null)
            {
                Marshal.ReleaseComObject(_functions);
                _functions = null;
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