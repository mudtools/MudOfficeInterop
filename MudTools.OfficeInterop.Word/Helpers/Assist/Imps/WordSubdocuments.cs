//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 子文档集合的封装实现类。
/// </summary>
internal class WordSubdocuments : IWordSubdocuments
{
    private MsWord.Subdocuments _subdocuments;
    private bool _disposedValue;

    internal WordSubdocuments(MsWord.Subdocuments subdocuments)
    {
        _subdocuments = subdocuments ?? throw new ArgumentNullException(nameof(subdocuments));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _subdocuments != null ? new WordApplication(_subdocuments.Application) : null;

    /// <inheritdoc/>
    public object Parent => _subdocuments?.Parent;

    /// <inheritdoc/>
    public int Count => _subdocuments?.Count ?? 0;

    /// <inheritdoc/>
    public IWordSubdocument this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comSubdocument = _subdocuments[index];
            return new WordSubdocument(comSubdocument);
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordSubdocument AddFromFile(string name, bool confirmConversions, bool readOnly, object passwordDocument,
                        string passwordTemplate, bool revert, string writePasswordDocument, string writePasswordTemplate)
    {
        if (name == null) throw new ArgumentNullException(nameof(name));

        try
        {
            var newSubdocument = _subdocuments.AddFromFile(name, confirmConversions, readOnly, passwordDocument,
                                                 passwordTemplate, revert, writePasswordDocument, writePasswordTemplate);
            return newSubdocument != null ? new WordSubdocument(newSubdocument) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加子文档。", ex);
        }
    }

    public IWordSubdocument AddFromRange(IWordRange range)
    {
        if (range == null) throw new ArgumentNullException(nameof(range));

        try
        {
            var newSubdocument = _subdocuments.AddFromRange(((WordRange)range)._range);
            return newSubdocument != null ? new WordSubdocument(newSubdocument) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加子文档。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete()
    {
        if (_subdocuments == null) return;

        _subdocuments?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _subdocuments?.Select();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _subdocuments != null)
        {
            Marshal.ReleaseComObject(_subdocuments);
            _subdocuments = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable 实现

    public IEnumerator<IWordSubdocument> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}