//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 列表段落集合的封装实现类。
/// </summary>
internal class WordListParagraphs : IWordListParagraphs
{
    private MsWord.ListParagraphs _listParagraphs;
    private bool _disposedValue;

    internal WordListParagraphs(MsWord.ListParagraphs listParagraphs)
    {
        _listParagraphs = listParagraphs ?? throw new ArgumentNullException(nameof(listParagraphs));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listParagraphs != null ? new WordApplication(_listParagraphs.Application) : null;

    /// <inheritdoc/>
    public object Parent => _listParagraphs?.Parent;

    /// <inheritdoc/>
    public int Count => _listParagraphs?.Count ?? 0;

    /// <inheritdoc/>
    public IWordParagraph this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comParagraph = _listParagraphs[index];
            return new WordParagraph(comParagraph);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listParagraphs != null)
        {
            Marshal.ReleaseComObject(_listParagraphs);
            _listParagraphs = null;
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

    public IEnumerator<IWordParagraph> GetEnumerator()
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