//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Comment 的封装实现类。
/// </summary>
internal class WordComment : IWordComment
{
    private MsWord.Comment _comment;
    private bool _disposedValue;

    internal WordComment(MsWord.Comment comment)
    {
        _comment = comment ?? throw new ArgumentNullException(nameof(comment));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _comment != null ? new WordApplication(_comment.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _comment?.Parent;

    /// <inheritdoc/>
    public int Index => _comment?.Index ?? 0;

    /// <inheritdoc/>
    public string Author
    {
        get => _comment?.Author ?? string.Empty;
        set
        {
            if (_comment != null && !string.IsNullOrEmpty(value))
            {
                _comment.Author = value;
            }
        }
    }

    /// <inheritdoc/>
    public string Initial
    {
        get => _comment?.Initial ?? string.Empty;
        set
        {
            if (_comment != null && !string.IsNullOrEmpty(value))
            {
                _comment.Initial = value;
            }
        }
    }

    /// <inheritdoc/>
    public IWordRange Range => _comment?.Reference != null ? new WordRange(_comment.Reference) : null;

    /// <inheritdoc/>
    public IWordRange CommentRange => _comment?.Scope != null ? new WordRange(_comment.Scope) : null;

    /// <inheritdoc/>
    public DateTime Date => _comment?.Date ?? DateTime.MinValue;

    /// <inheritdoc/>
    public int CharactersCount => _comment?.Scope?.Characters?.Count ?? 0;

    /// <inheritdoc/>
    public int WordsCount => _comment?.Scope?.Words?.Count ?? 0;

    /// <inheritdoc/>
    public int SentencesCount => _comment?.Scope?.Sentences?.Count ?? 0;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _comment?.Reference?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _comment?.Delete();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _comment?.Scope?.Copy();
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _comment?.Scope?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_comment?.Scope != null)
        {
            _comment.Scope.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public void AppendText(string text)
    {
        if (_comment?.Scope != null && !string.IsNullOrEmpty(text))
        {
            var endRange = _comment.Scope.Duplicate;
            endRange.Collapse(MsWord.WdCollapseDirection.wdCollapseEnd);
            endRange.Text = text;
        }
    }

    /// <inheritdoc/>
    public bool ContainsText(string text, bool matchCase = false)
    {
        if (string.IsNullOrEmpty(text)) return false;

        try
        {
            string commentText = GetText();
            return matchCase ?
                commentText.Contains(text) :
                commentText.ToLower().Contains(text.ToLower());
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public int ReplaceText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (_comment?.Scope == null || string.IsNullOrEmpty(findText)) return 0;

        int replaceCount = 0;
        try
        {
            var findObj = _comment.Scope.Find;
            findObj.ClearFormatting();
            findObj.Replacement.ClearFormatting();
            findObj.Text = findText;
            findObj.Replacement.Text = replaceText ?? string.Empty;
            findObj.Forward = true;
            findObj.Wrap = MsWord.WdFindWrap.wdFindStop;
            findObj.Format = false;
            findObj.MatchCase = matchCase;
            findObj.MatchWholeWord = matchWholeWord;
            findObj.MatchWildcards = false;
            findObj.MatchSoundsLike = false;
            findObj.MatchAllWordForms = false;

            // 替换所有匹配项
            while (findObj.Execute(Replace: MsWord.WdReplace.wdReplaceAll))
            {
                replaceCount++;
            }
        }
        catch
        {
            // 替换失败返回 0
        }

        return replaceCount;
    }

    /// <inheritdoc/>
    public string GetReferenceText()
    {
        return _comment?.Reference?.Text ?? string.Empty;
    }

    /// <inheritdoc/>
    public IWordRange GetReferenceRange()
    {
        return _comment?.Reference != null ? new WordRange(_comment.Reference) : null;
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_comment != null)
            {
                Marshal.ReleaseComObject(_comment);
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