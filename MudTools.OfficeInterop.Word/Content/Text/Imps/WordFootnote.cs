//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Footnote 的封装实现类。
/// </summary>
internal class WordFootnote : IWordFootnote
{
    private MsWord.Footnote _footnote;
    private bool _disposedValue;

    internal WordFootnote(MsWord.Footnote footnote)
    {
        _footnote = footnote ?? throw new ArgumentNullException(nameof(footnote));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _footnote != null ? new WordApplication(_footnote.Application) : null;

    /// <inheritdoc/>
    public object Parent => _footnote?.Parent;

    /// <inheritdoc/>
    public int Index => _footnote?.Index ?? 0;

    /// <inheritdoc/>
    public IWordRange Reference => _footnote?.Reference != null ? new WordRange(_footnote.Reference) : null;

    /// <inheritdoc/>
    public IWordRange Range => _footnote?.Range != null ? new WordRange(_footnote.Range) : null;

    /// <inheritdoc/>
    public string Number => _footnote?.Reference?.Text ?? string.Empty;


    /// <inheritdoc/>
    public IWordFont Font
    {
        get => _footnote?.Range?.Font != null ? new WordFont(_footnote.Range.Font) : null;
    }

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat
    {
        get => _footnote?.Range?.ParagraphFormat != null ? new WordParagraphFormat(_footnote.Range.ParagraphFormat) : null;
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _footnote?.Range?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _footnote?.Delete();
    }

    /// <inheritdoc/>
    public IWordFootnote Copy()
    {
        if (_footnote == null) return null;

        _footnote.Range.Copy();
        return this; // 返回当前脚注作为"复制"的标识
    }

    /// <inheritdoc/>
    public void Update()
    {
        if (_footnote == null) return;

        var parentDocument = _footnote.Application.ActiveDocument;
        parentDocument.Fields.Update();
    }

    /// <inheritdoc/>
    public IWordRange GetReferenceRange()
    {
        return _footnote?.Reference != null ? new WordRange(_footnote.Reference) : null;
    }

    /// <inheritdoc/>
    public IWordRange GetContentRange()
    {
        return _footnote?.Range != null ? new WordRange(_footnote.Range) : null;
    }

    /// <inheritdoc/>
    public void ModifyText(string newText)
    {
        if (_footnote?.Range != null && newText != null)
        {
            _footnote.Range.Text = newText;
        }
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _footnote?.Range?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_footnote?.Range != null)
        {
            _footnote.Range.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public bool ContainsText(string text, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(text)) return false;

        try
        {
            string footnoteText = GetText();
            return matchCase ?
                footnoteText.Contains(text) :
                footnoteText.ToLower().Contains(text.ToLower());
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
            if (_footnote != null)
            {
                Marshal.ReleaseComObject(_footnote);
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