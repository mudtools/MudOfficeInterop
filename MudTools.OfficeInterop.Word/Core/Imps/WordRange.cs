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
    /// 初始化 <see cref="WordRange"/> 类的新实例。
    /// </summary>
    /// <param name="range">要封装的原始 COM Range 对象。</param>
    internal WordRange(MsWord.Range range)
    {
        _range = range ?? throw new ArgumentNullException(nameof(range));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication? Application => _range != null ? new WordApplication(_range.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _range?.Parent;

    /// <inheritdoc/>
    public int Start
    {
        get => _range?.Start ?? 0;
        set { if (_range != null) _range.Start = value; }
    }

    /// <inheritdoc/>
    public int End
    {
        get => _range?.End ?? 0;
        set { if (_range != null) _range.End = value; }
    }

    /// <inheritdoc/>
    public string Text
    {
        get => _range?.Text ?? string.Empty;
        set { if (_range != null) _range.Text = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public IWordRange? Duplicate => _range?.Duplicate != null ? new WordRange(_range.Duplicate) : null;

    /// <inheritdoc/>
    public IWordDocument? Document => _range?.Document != null ? new WordDocument(_range.Document) : null;

    /// <inheritdoc/>
    public WdStoryType? StoryType => _range?.StoryType != null ? (WdStoryType)(int)_range?.StoryType : WdStoryType.wdMainTextStory;

    /// <inheritdoc/>
    public int StoryLength => _range?.StoryLength ?? 0;

    /// <inheritdoc/>
    public IWordRange? NextStoryRange => _range?.NextStoryRange != null ? new WordRange(_range.NextStoryRange) : null;

    #endregion

    #region 格式化属性实现 (Formatting Properties Implementation)

    /// <inheritdoc/>
    public IWordFont? Font => _range?.Font != null ? new WordFont(_range.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat? ParagraphFormat => _range?.ParagraphFormat != null ? new WordParagraphFormat(_range.ParagraphFormat) : null;

    /// <inheritdoc/>
    public object? Style
    {
        get => _range?.get_Style();
        set { _range?.set_Style(value); }
    }

    /// <inheritdoc/>
    public int Bold
    {
        get => _range?.Bold ?? 0;
        set { if (_range != null) _range.Bold = value; }
    }

    /// <inheritdoc/>
    public int Italic
    {
        get => _range?.Italic ?? 0;
        set { if (_range != null) _range.Italic = value; }
    }

    /// <inheritdoc/>
    public WdUnderline Underline
    {
        get => _range?.Underline != null ? (WdUnderline)(int)_range?.Underline : WdUnderline.wdUnderlineNone;
        set
        {
            if (_range != null) _range.Underline = (MsWord.WdUnderline)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColorIndex HighlightColorIndex
    {
        get => _range?.HighlightColorIndex != null ? (WdColorIndex)(int)_range?.HighlightColorIndex : WdColorIndex.wdRed;
        set
        {
            if (_range != null) _range.HighlightColorIndex = (MsWord.WdColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdCharacterCase Case
    {
        get => _range?.Case != null ? (WdCharacterCase)(int)_range?.Case : WdCharacterCase.wdLowerCase;
        set
        {
            if (_range != null) _range.Case = (MsWord.WdCharacterCase)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordShading? Shading => _range?.Shading != null ? new WordShading(_range.Shading) : null;

    /// <inheritdoc/>
    public IWordListFormat? ListFormat => _range?.ListFormat != null ? new WordListFormat(_range.ListFormat) : null;
    /// <inheritdoc/>
    public IWordPageSetup? PageSetup
    {
        get => _range?.PageSetup != null ? new WordPageSetup(_range.PageSetup) : null;
        set { if (_range != null) _range.PageSetup = ((WordPageSetup)value)._pageSetup; }
    }

    #endregion

    #region 集合属性实现 (Collection Properties Implementation - 第一部分)

    /// <inheritdoc/>
    public IWordParagraphs? Paragraphs => _range?.Paragraphs != null ? new WordParagraphs(_range.Paragraphs) : null;

    /// <inheritdoc/>
    public IWordSentences? Sentences => _range?.Sentences != null ? new WordSentences(_range.Sentences) : null;

    /// <inheritdoc/>
    public IWordWords? Words => _range?.Words != null ? new WordWords(_range.Words) : null;

    /// <inheritdoc/>
    public IWordCharacters? Characters => _range?.Characters != null ? new WordCharacters(_range.Characters) : null;

    /// <inheritdoc/>
    public IWordTables? Tables => _range?.Tables != null ? new WordTables(_range.Tables) : null;

    /// <inheritdoc/>
    public IWordBookmarks? Bookmarks => _range?.Bookmarks != null ? new WordBookmarks(_range.Bookmarks) : null;

    /// <inheritdoc/>
    public IWordFields? Fields => _range?.Fields != null ? new WordFields(_range.Fields) : null;

    /// <inheritdoc/>
    public IWordHyperlinks? Hyperlinks => _range?.Hyperlinks != null ? new WordHyperlinks(_range.Hyperlinks) : null;

    /// <inheritdoc/>
    public IWordFormFields? FormFields => _range?.FormFields != null ? new WordFormFields(_range.FormFields) : null;

    /// <inheritdoc/>
    public IWordRevisions? Revisions => _range?.Revisions != null ? new WordRevisions(_range.Revisions) : null;

    /// <inheritdoc/>
    public IWordComments? Comments => _range?.Comments != null ? new WordComments(_range.Comments) : null;

    /// <inheritdoc/>
    public IWordFootnotes? Footnotes => _range?.Footnotes != null ? new WordFootnotes(_range.Footnotes) : null;

    /// <inheritdoc/>
    public IWordEndnotes? Endnotes => _range?.Endnotes != null ? new WordEndnotes(_range.Endnotes) : null;

    #endregion

    #region 状态和工具属性实现 (State & Utility Properties Implementation)

    /// <inheritdoc/>
    public bool SpellingChecked
    {
        get => _range?.SpellingChecked ?? false;
        set { if (_range != null) _range.SpellingChecked = value; }
    }

    /// <inheritdoc/>
    public bool GrammarChecked
    {
        get => _range?.GrammarChecked ?? false;
        set { if (_range != null) _range.GrammarChecked = value; }
    }

    /// <inheritdoc/>
    public bool NoProofing
    {
        get => _range?.NoProofing != null && _range?.NoProofing == 1;
        set { if (_range != null) _range.NoProofing = value ? 1 : 0; }
    }

    /// <inheritdoc/>
    public IWordFind? Find => _range?.Find != null ? new WordFind(_range.Find) : null;

    /// <inheritdoc/>
    public IWordRange? FormattedText
    {
        get => _range?.FormattedText != null ? new WordRange(_range.FormattedText) : null;
        set
        {
            if (_range != null && value is WordRange sourceRange)
            {
                _range.FormattedText = sourceRange._range;
            }
        }
    }

    #endregion

    #region 基本方法实现 (Basic Methods Implementation)

    /// <inheritdoc/>
    public void Select()
    {
        _range?.Select();
    }

    /// <inheritdoc/>
    public void SetRange(int start, int end)
    {
        _range?.SetRange(start, end);
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
    public void Delete()
    {
        _range?.Delete();
    }

    /// <inheritdoc/>
    public void CopyAsText(IWordRange source)
    {
        if (_range == null || source == null) return;
        try
        {
            var sourceRange = source as WordRange;
            if (sourceRange?._range != null)
            {
                _range.FormattedText = sourceRange._range.FormattedText;
            }
        }
        catch (COMException ex)
        {
            System.Diagnostics.Debug.WriteLine($"CopyAsText failed: {ex.Message}");
        }
    }

    /// <inheritdoc/>
    public object Information(WdInformation type)
    {
        return _range?.Information[(MsWord.WdInformation)(int)type];
    }

    /// <inheritdoc/>
    public bool FindAndReplace(object findText, object replaceWith, MsWord.WdReplace replace)
    {
        if (_range?.Find == null) return false;

        _range.Find.ClearFormatting();
        _range.Find.Text = findText?.ToString() ?? string.Empty;
        _range.Find.Replacement.ClearFormatting();
        _range.Find.Replacement.Text = replaceWith?.ToString() ?? string.Empty;

        return _range.Find.Execute(ref findText);
    }

    #endregion


    #region 更多集合属性实现 (More Collection Properties Implementation)

    /// <inheritdoc/>
    public IWordShapeRange? ShapeRange
    {
        get
        {
            try { return _range?.ShapeRange != null ? new WordShapeRange(_range.ShapeRange) : null; }
            catch (COMException) { return null; }
        }
    }

    /// <inheritdoc/>
    public IWordInlineShapes? InlineShapes => _range?.InlineShapes != null ? new WordInlineShapes(_range.InlineShapes) : null;

    /// <inheritdoc/>
    public IWordBorders? Borders => _range?.Borders != null ? new WordBorders(_range.Borders) : null;

    /// <inheritdoc/>
    public IWordListParagraphs? ListParagraphs => _range?.ListParagraphs != null ? new WordListParagraphs(_range.ListParagraphs) : null;

    /// <inheritdoc/>
    public IWordReadabilityStatistics? ReadabilityStatistics => _range?.ReadabilityStatistics != null ? new WordReadabilityStatistics(_range.ReadabilityStatistics) : null;

    /// <inheritdoc/>
    public IWordProofreadingErrors? SpellingErrors => _range?.SpellingErrors != null ? new WordProofreadingErrors(_range.SpellingErrors) : null;

    /// <inheritdoc/>
    public IWordProofreadingErrors? GrammaticalErrors => _range?.GrammaticalErrors != null ? new WordProofreadingErrors(_range.GrammaticalErrors) : null;

    /// <inheritdoc/>
    public IWordSubdocuments? Subdocuments => _range?.Subdocuments != null ? new WordSubdocuments(_range.Subdocuments) : null;

    /// <inheritdoc/>
    public IWordContentControls? ContentControls => _range?.ContentControls != null ? new WordContentControls(_range.ContentControls) : null;

    public IWordConflicts? Conflicts => _range?.Conflicts != null ? new WordConflicts(_range.Conflicts) : null;

    public IWordEditors? Editors => _range?.Editors != null ? new WordEditors(_range.Editors) : null;


    #endregion

    #region 更多格式化属性实现 (More Formatting Properties Implementation)

    /// <inheritdoc/>
    public int BoldBi
    {
        get => _range?.BoldBi ?? 0;
        set { if (_range != null) _range.BoldBi = value; }
    }

    /// <inheritdoc/>
    public int ItalicBi
    {
        get => _range?.ItalicBi ?? 0;
        set { if (_range != null) _range.ItalicBi = value; }
    }

    /// <inheritdoc/>
    public WdEmphasisMark EmphasisMark
    {
        get => _range?.EmphasisMark != null ? (WdEmphasisMark)(int)_range?.EmphasisMark : WdEmphasisMark.wdEmphasisMarkNone;
        set
        {
            if (_range != null) _range.EmphasisMark = (MsWord.WdEmphasisMark)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdCharacterWidth CharacterWidth
    {
        get => _range?.CharacterWidth != null ? (WdCharacterWidth)(int)_range?.CharacterWidth : WdCharacterWidth.wdWidthHalfWidth;
        set
        {
            if (_range != null) _range.CharacterWidth = (MsWord.WdCharacterWidth)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdHorizontalInVerticalType HorizontalInVertical
    {
        get => _range?.HorizontalInVertical != null ? (WdHorizontalInVerticalType)(int)_range?.HorizontalInVertical : WdHorizontalInVerticalType.wdHorizontalInVerticalNone;
        set
        {
            if (_range != null) _range.HorizontalInVertical = (MsWord.WdHorizontalInVerticalType)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdTextOrientation Orientation
    {
        get => _range?.Orientation != null ? (WdTextOrientation)(int)_range?.Orientation : WdTextOrientation.wdTextOrientationDownward;
        set
        {
            if (_range != null) _range.Orientation = (MsWord.WdTextOrientation)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdTwoLinesInOneType TwoLinesInOne
    {
        get => _range?.TwoLinesInOne != null ? (WdTwoLinesInOneType)(int)_range?.TwoLinesInOne : WdTwoLinesInOneType.wdTwoLinesInOneNone;
        set
        {
            if (_range != null) _range.TwoLinesInOne = (MsWord.WdTwoLinesInOneType)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageID
    {
        get => _range?.LanguageID != null ? (WdLanguageID)(int)_range?.LanguageID : WdLanguageID.wdSimplifiedChinese;
        set
        {
            if (_range != null) _range.LanguageID = (MsWord.WdLanguageID)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageIDFarEast
    {
        get => _range?.LanguageIDFarEast != null ? (WdLanguageID)(int)_range?.LanguageIDFarEast : WdLanguageID.wdSimplifiedChinese;
        set
        {
            if (_range != null) _range.LanguageIDFarEast = (MsWord.WdLanguageID)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageIDOther
    {
        get => _range?.LanguageIDOther != null ? (WdLanguageID)(int)_range?.LanguageIDOther : WdLanguageID.wdSimplifiedChinese;
        set
        {
            if (_range != null) _range.LanguageIDOther = (MsWord.WdLanguageID)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool LanguageDetected
    {
        get => _range?.LanguageDetected ?? false;
        set { if (_range != null) _range.LanguageDetected = value; }
    }


    /// <inheritdoc/>
    public bool DisableCharacterSpaceGrid
    {
        get => _range?.DisableCharacterSpaceGrid ?? false;
        set { if (_range != null) _range.DisableCharacterSpaceGrid = value; }
    }

    /// <inheritdoc/>
    public string ID
    {
        get => _range?.ID ?? string.Empty;
        set { if (_range != null) _range.ID = value ?? string.Empty; }
    }

    #endregion

    #region 更多方法实现 (More Methods Implementation)

    /// <inheritdoc/>
    public void CheckSpelling()
    {
        _range?.CheckSpelling();
    }

    /// <inheritdoc/>
    public void CheckGrammar()
    {
        _range?.CheckGrammar();
    }

    /// <inheritdoc/>
    public IWordBookmark? Bookmark(string name)
    {
        if (_range?.Bookmarks == null || string.IsNullOrWhiteSpace(name)) return null;
        var bookmark = _range.Bookmarks.Add(name, _range);
        return bookmark != null ? new WordBookmark(bookmark) : null;
    }

    /// <inheritdoc/>
    public IWordHyperlink? Hyperlink(object address, object subAddress, object screenTip, object textToDisplay, object target)
    {
        if (_range?.Hyperlinks == null) return null;
        var hyperlink = _range.Hyperlinks.Add(_range, address, subAddress, screenTip, textToDisplay, target);
        return hyperlink != null ? new WordHyperlink(hyperlink) : null;
    }

    /// <inheritdoc/>
    public void CellsMerge()
    {
        _range?.Cells?.Merge();
    }

    /// <inheritdoc/>
    public void CellsSplit(int numRows, int numColumns, bool mergeBeforeSplit)
    {
        _range?.Cells?.Split(numRows, numColumns, mergeBeforeSplit);
    }

    /// <inheritdoc/>
    public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object ascending)
    {
        _range?.Sort(ref excludeHeader, ref fieldNumber, ref sortFieldType, ref ascending);
    }

    /// <inheritdoc/>
    public void SaveAsText(string fileName)
    {
        if (_range == null || string.IsNullOrWhiteSpace(fileName)) return;
        try
        {
            File.WriteAllText(fileName, this.Text);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"SaveAsText failed: {ex.Message}");
            // 可以选择抛出异常
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _range != null)
        {
            Marshal.ReleaseComObject(_range);
            _range = null;
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