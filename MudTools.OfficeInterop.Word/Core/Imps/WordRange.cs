//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Range 的实现类。
/// </summary>
internal class WordRange : IWordRange
{
    internal MsWord.Range _range;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="range">原始 COM Range 对象。</param>
    internal WordRange(MsWord.Range range)
    {
        _range = range ?? throw new ArgumentNullException(nameof(range));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _range != null ? new WordApplication(_range.Application) : null;


    /// <inheritdoc/>
    public string Text
    {
        get => _range?.Text ?? string.Empty;
        set
        {
            if (_range != null)
                _range.Text = value;
        }
    }

    /// <inheritdoc/>
    public int Start
    {
        get => _range?.Start ?? 0;
        set
        {
            if (_range != null)
                _range.Start = value;
        }
    }

    /// <inheritdoc/>
    public int End
    {
        get => _range?.End ?? 0;
        set
        {
            if (_range != null)
                _range.End = value;
        }
    }

    /// <inheritdoc/>
    public IWordFont? Font => _range?.Font != null ? new WordFont(_range.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat? ParagraphFormat =>
        _range?.ParagraphFormat != null ? new WordParagraphFormat(_range.ParagraphFormat) : null;

    /// <inheritdoc/>
    public IWordCharacters? Characters =>
        _range?.Characters != null ? new WordCharacters(_range.Characters) : null;


    /// <inheritdoc/>
    public int CharactersCount => _range?.Characters?.Count ?? 0;

    /// <inheritdoc/>
    public int WordsCount => _range?.Words?.Count ?? 0;

    /// <inheritdoc/>
    public int ParagraphsCount => _range?.Paragraphs?.Count ?? 0;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void InsertAfter(string text)
    {
        if (_range == null) return;

        _range.InsertAfter(text);
    }

    /// <inheritdoc/>
    public void InsertBefore(string text)
    {
        if (_range == null) return;

        _range.InsertBefore(text);
    }

    /// <inheritdoc/>
    public void InsertParagraphAfter()
    {
        _range?.InsertParagraphAfter();
    }

    /// <inheritdoc/>
    public void InsertParagraphBefore()
    {
        _range?.InsertParagraphBefore();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _range?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _range?.Select();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _range?.Copy();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _range?.Paste();
    }

    /// <inheritdoc/>
    public bool FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (_range == null || string.IsNullOrEmpty(findText))
            return false;

        var find = _range.Document.Range().Find;
        find.ClearFormatting();
        find.Text = findText;
        find.Replacement.ClearFormatting();
        find.Replacement.Text = replaceText;
        find.Forward = true;
        find.Wrap = MsWord.WdFindWrap.wdFindContinue;
        find.Format = false;
        find.MatchCase = matchCase;
        find.MatchWholeWord = matchWholeWord;
        find.MatchWildcards = false;
        find.MatchSoundsLike = false;
        find.MatchAllWordForms = false;

        return find.Execute();
    }

    /// <inheritdoc/>
    public IWordRange GetCharacter(int index)
    {
        if (_range?.Characters == null || index < 1 || index > CharactersCount)
            return null;

        var charRange = _range.Characters[index];
        return new WordRange(charRange);
    }

    /// <inheritdoc/>
    public IWordRange GetWord(int index)
    {
        if (_range?.Words == null || index < 1 || index > WordsCount)
            return null;

        var wordRange = _range.Words[index];
        return new WordRange(wordRange);
    }

    /// <inheritdoc/>
    public IWordRange GetParagraph(int index)
    {
        if (_range?.Paragraphs == null || index < 1 || index > ParagraphsCount)
            return null;

        var paraRange = _range.Paragraphs[index].Range;
        return new WordRange(paraRange);
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放字体和段落格式对象
            if (_range?.Font != null)
            {
                Marshal.ReleaseComObject(_range.Font);
            }
            if (_range?.ParagraphFormat != null)
            {
                Marshal.ReleaseComObject(_range.ParagraphFormat);
            }
            // 释放范围对象本身
            if (_range != null)
            {
                Marshal.ReleaseComObject(_range);
                _range = null;
            }
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    #endregion
}