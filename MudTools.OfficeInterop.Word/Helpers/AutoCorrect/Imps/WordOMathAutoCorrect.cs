//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordOMathAutoCorrect : IWordOMathAutoCorrect
{
    private MsWord.OMathAutoCorrect _autoCorrect;
    private bool _disposedValue;

    internal WordOMathAutoCorrect(MsWord.OMathAutoCorrect autoCorrect)
    {
        _autoCorrect = autoCorrect ?? throw new ArgumentNullException(nameof(autoCorrect));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _autoCorrect?.Application != null ? new WordApplication(_autoCorrect.Application) : null;

    /// <summary>
    /// 获取代表指定对象的父对象的对象。
    /// </summary>
    public object Parent => _autoCorrect?.Parent;

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    public int Creator => _autoCorrect?.Creator ?? 0;

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值代表是否将键入的普通文本自动更正为专业格式的数学符号。
    /// </summary>
    public bool UseOutsideOMath
    {
        get => _autoCorrect?.UseOutsideOMath ?? false;
        set
        {
            if (_autoCorrect != null)
                _autoCorrect.UseOutsideOMath = value;
        }
    }

    private IWordOMathAutoCorrectEntries _entries;
    /// <summary>
    /// 获取数学自动更正条目的集合。
    /// </summary>
    public IWordOMathAutoCorrectEntries Entries
    {
        get
        {
            if (_entries == null && _autoCorrect?.Entries != null)
            {
                _entries = new WordOMathAutoCorrectEntries(_autoCorrect.Entries);
            }
            return _entries;
        }
    }

    private IWordOMathRecognizedFunctions _functions;
    /// <summary>
    /// 获取数学识别函数的集合。
    /// </summary>
    public IWordOMathRecognizedFunctions Functions
    {
        get
        {
            if (_functions == null && _autoCorrect?.Functions != null)
            {
                _functions = new WordOMathRecognizedFunctions(_autoCorrect.Functions);
            }
            return _functions;
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放子集合
            _entries?.Dispose();
            _functions?.Dispose();

            // 释放主对象
            if (_autoCorrect != null)
            {
                Marshal.ReleaseComObject(_autoCorrect);
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