//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word.Endnote 的封装实现类。
/// </summary>
internal class WordEndnote : IWordEndnote
{
    private MsWord.Endnote _endnote;
    private bool _disposedValue;

    internal WordEndnote(MsWord.Endnote endnote)
    {
        _endnote = endnote ?? throw new ArgumentNullException(nameof(endnote));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _endnote != null ? new WordApplication(_endnote.Application) : null;

    /// <inheritdoc/>
    public object Parent => _endnote?.Parent;

    /// <inheritdoc/>
    public int Index => _endnote?.Index ?? 0;

    /// <inheritdoc/>
    public IWordRange Reference => _endnote?.Reference != null ? new WordRange(_endnote.Reference) : null;

    /// <inheritdoc/>
    public IWordRange Range => _endnote?.Range != null ? new WordRange(_endnote.Range) : null;

    /// <inheritdoc/>
    public string Number => _endnote?.Reference?.Text ?? string.Empty;

    /// <inheritdoc/>
    public IWordFont Font
    {
        get => _endnote?.Range?.Font != null ? new WordFont(_endnote.Range.Font) : null;
    }

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat
    {
        get => _endnote?.Range?.ParagraphFormat != null ? new WordParagraphFormat(_endnote.Range.ParagraphFormat) : null;
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _endnote?.Range?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _endnote?.Delete();
    }

    /// <inheritdoc/>
    public IWordEndnote Copy()
    {
        if (_endnote == null) return null;

        _endnote.Range.Copy();
        return this; // 返回当前尾注作为"复制"的标识
    }

    /// <inheritdoc/>
    public void Update()
    {
        if (_endnote == null) return;

        // 更新通常通过更新域来实现
        var parentDocument = _endnote.Application.ActiveDocument;
        parentDocument.Fields.Update();
    }

    /// <inheritdoc/>
    public IWordRange GetReferenceRange()
    {
        return _endnote?.Reference != null ? new WordRange(_endnote.Reference) : null;
    }

    /// <inheritdoc/>
    public IWordRange GetContentRange()
    {
        return _endnote?.Range != null ? new WordRange(_endnote.Range) : null;
    }

    /// <inheritdoc/>
    public void ModifyText(string newText)
    {
        if (_endnote?.Range != null && newText != null)
        {
            _endnote.Range.Text = newText;
        }
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _endnote?.Range?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_endnote?.Range != null)
        {
            _endnote.Range.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public bool ContainsText(string text, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(text)) return false;

        try
        {
            string endnoteText = GetText();
            return matchCase ?
                endnoteText.Contains(text) :
                endnoteText.ToLower().Contains(text.ToLower());
        }
        catch
        {
            return false;
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_endnote != null)
            {
                Marshal.ReleaseComObject(_endnote);
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