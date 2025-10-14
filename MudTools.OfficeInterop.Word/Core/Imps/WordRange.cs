//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Range 的实现类。
/// </summary>
internal class WordRange : IWordRange
{
    internal MsWord.Range _range;
    private bool _disposedValue;
    private IWordRange? _duplicate;
    private IWordDocument? _document;
    private IWordFont? _font;
    private IWordParagraphFormat? _paragraphFormat;
    private IWordShading? _shading;
    private IWordListFormat? _listFormat;
    private IWordPageSetup? _pageSetup;
    private IWordParagraphs? _paragraphs;
    private IWordSentences? _sentences;
    private IWordWords? _words;
    private IWordCharacters? _characters;
    private IWordTables? _tables;
    private IWordBookmarks? _bookmarks;
    private IWordFields? _fields;
    private IWordHyperlinks? _hyperlinks;
    private IWordFormFields? _formFields;
    private IWordRevisions? _revisions;
    private IWordComments? _comments;
    private IWordFootnotes? _footnotes;
    private IWordEndnotes? _endnotes;
    private IWordFind? _find;
    private IWordRange? _formattedText;
    private IWordShapeRange? _shapeRangeField;  // 修复字段名冲突
    private IWordInlineShapes? _inlineShapes;
    private IWordBorders? _borders;
    private IWordListParagraphs? _listParagraphs;
    private IWordReadabilityStatistics? _readabilityStatistics;
    private IWordProofreadingErrors? _spellingErrors;
    private IWordProofreadingErrors? _grammaticalErrors;
    private IWordSubdocuments? _subdocuments;
    private IWordContentControls? _contentControls;
    private IWordConflicts? _conflicts;
    private IWordEditors? _editors;
    private static readonly ILog log = LogManager.GetLogger(typeof(WordRange));

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
    public IWordRange? Duplicate
    {
        get
        {
            if (_range?.Duplicate != null)
            {
                _duplicate ??= new WordRange(_range.Duplicate);
                return _duplicate;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordDocument? Document
    {
        get
        {
            if (_range?.Document != null)
            {
                _document ??= new WordDocument(_range.Document);
                return _document;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public WdStoryType? StoryType => _range?.StoryType != null ? (WdStoryType)(int)_range?.StoryType : WdStoryType.wdMainTextStory;

    /// <inheritdoc/>
    public int StoryLength => _range?.StoryLength ?? 0;

    /// <inheritdoc/>
    public IWordRange? NextStoryRange => _range?.NextStoryRange != null ? new WordRange(_range.NextStoryRange) : null;

    #endregion

    #region 格式化属性实现 (Formatting Properties Implementation)

    /// <inheritdoc/>
    public IWordFont? Font
    {
        get
        {
            if (_range?.Font != null)
            {
                _font ??= new WordFont(_range.Font);
                return _font;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordParagraphFormat? ParagraphFormat
    {
        get
        {
            if (_range?.ParagraphFormat != null)
            {
                _paragraphFormat ??= new WordParagraphFormat(_range.ParagraphFormat);
                return _paragraphFormat;
            }
            return null;
        }
    }

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
        get => _range?.Underline != null ? _range.Underline.EnumConvert(WdUnderline.wdUnderlineNone) : WdUnderline.wdUnderlineNone;
        set
        {
            if (_range != null) _range.Underline = value.EnumConvert(MsWord.WdUnderline.wdUnderlineNone);
        }
    }

    /// <inheritdoc/>
    public WdColorIndex HighlightColorIndex
    {
        get => _range?.HighlightColorIndex != null ? _range.HighlightColorIndex.EnumConvert(WdColorIndex.wdAuto) : WdColorIndex.wdAuto;
        set
        {
            if (_range != null) _range.HighlightColorIndex = value.EnumConvert(MsWord.WdColorIndex.wdAuto);
        }
    }

    /// <inheritdoc/>
    public WdCharacterCase Case
    {
        get => _range?.Case != null ? _range.Case.EnumConvert(WdCharacterCase.wdNextCase) : WdCharacterCase.wdNextCase;
        set
        {
            if (_range != null) _range.Case = value.EnumConvert(MsWord.WdCharacterCase.wdNextCase);
        }
    }

    /// <inheritdoc/>
    public IWordShading? Shading
    {
        get
        {
            if (_range?.Shading != null)
            {
                _shading ??= new WordShading(_range.Shading);
                return _shading;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordListFormat? ListFormat
    {
        get
        {
            if (_range?.ListFormat != null)
            {
                _listFormat ??= new WordListFormat(_range.ListFormat);
                return _listFormat;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordPageSetup? PageSetup
    {
        get
        {
            if (_range?.PageSetup != null)
            {
                _pageSetup ??= new WordPageSetup(_range.PageSetup);
                return _pageSetup;
            }
            return null;
        }
        set { if (_range != null) _range.PageSetup = ((WordPageSetup)value)._pageSetup; }
    }

    #endregion

    #region 集合属性实现 (Collection Properties Implementation - 第一部分)

    /// <inheritdoc/>
    public IWordParagraphs? Paragraphs
    {
        get
        {
            if (_range?.Paragraphs != null)
            {
                _paragraphs ??= new WordParagraphs(_range.Paragraphs);
                return _paragraphs;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordSentences? Sentences
    {
        get
        {
            if (_range?.Sentences != null)
            {
                _sentences ??= new WordSentences(_range.Sentences);
                return _sentences;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordWords? Words
    {
        get
        {
            if (_range?.Words != null)
            {
                _words ??= new WordWords(_range.Words);
                return _words;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordCharacters? Characters
    {
        get
        {
            if (_range?.Characters != null)
            {
                _characters ??= new WordCharacters(_range.Characters);
                return _characters;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordTables? Tables
    {
        get
        {
            if (_range?.Tables != null)
            {
                _tables ??= new WordTables(_range.Tables);
                return _tables;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordBookmarks? Bookmarks
    {
        get
        {
            if (_range?.Bookmarks != null)
            {
                _bookmarks ??= new WordBookmarks(_range.Bookmarks);
                return _bookmarks;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordFields? Fields
    {
        get
        {
            if (_range?.Fields != null)
            {
                _fields ??= new WordFields(_range.Fields);
                return _fields;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordHyperlinks? Hyperlinks
    {
        get
        {
            if (_range?.Hyperlinks != null)
            {
                _hyperlinks ??= new WordHyperlinks(_range.Hyperlinks);
                return _hyperlinks;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordFormFields? FormFields
    {
        get
        {
            if (_range?.FormFields != null)
            {
                _formFields ??= new WordFormFields(_range.FormFields);
                return _formFields;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordRevisions? Revisions
    {
        get
        {
            if (_range?.Revisions != null)
            {
                _revisions ??= new WordRevisions(_range.Revisions);
                return _revisions;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordComments? Comments
    {
        get
        {
            if (_range?.Comments != null)
            {
                _comments ??= new WordComments(_range.Comments);
                return _comments;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordFootnotes? Footnotes
    {
        get
        {
            if (_range?.Footnotes != null)
            {
                _footnotes ??= new WordFootnotes(_range.Footnotes);
                return _footnotes;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordEndnotes? Endnotes
    {
        get
        {
            if (_range?.Endnotes != null)
            {
                _endnotes ??= new WordEndnotes(_range.Endnotes);
                return _endnotes;
            }
            return null;
        }
    }

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
    public IWordFind? Find
    {
        get
        {
            if (_range?.Find != null)
            {
                _find ??= new WordFind(_range.Find);
                return _find;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordRange? FormattedText
    {
        get
        {
            if (_range?.FormattedText != null)
            {
                _formattedText ??= new WordRange(_range.FormattedText);
                return _formattedText;
            }
            return null;
        }
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
    public void Cut()
    {
        _range?.Cut();
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

    public void Collapse(WdCollapseDirection Direction)
    {
        _range?.Collapse(Direction.EnumConvert(WdCollapseDirection.wdCollapseStart));
    }

    public void Collapse()
    {
        _range?.Collapse();
    }


    public void InsertAfter(string text)
    {
        _range?.InsertAfter(text ?? string.Empty);
    }

    public void InsertBefore(string text)
    {
        _range?.InsertBefore(text ?? string.Empty);
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
            log.Error($"CopyAsText failed: {ex.Message}", ex);
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
            if (_range?.ShapeRange != null)
            {
                _shapeRangeField ??= new WordShapeRange(_range.ShapeRange);
                return _shapeRangeField;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordInlineShapes? InlineShapes
    {
        get
        {
            if (_range?.InlineShapes != null)
            {
                _inlineShapes ??= new WordInlineShapes(_range.InlineShapes);
                return _inlineShapes;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordBorders? Borders
    {
        get
        {
            if (_range?.Borders != null)
            {
                _borders ??= new WordBorders(_range.Borders);
                return _borders;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordListParagraphs? ListParagraphs
    {
        get
        {
            if (_range?.ListParagraphs != null)
            {
                _listParagraphs ??= new WordListParagraphs(_range.ListParagraphs);
                return _listParagraphs;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordReadabilityStatistics? ReadabilityStatistics
    {
        get
        {
            if (_range?.ReadabilityStatistics != null)
            {
                _readabilityStatistics ??= new WordReadabilityStatistics(_range.ReadabilityStatistics);
                return _readabilityStatistics;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordProofreadingErrors? SpellingErrors
    {
        get
        {
            if (_range?.SpellingErrors != null)
            {
                _spellingErrors ??= new WordProofreadingErrors(_range.SpellingErrors);
                return _spellingErrors;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordProofreadingErrors? GrammaticalErrors
    {
        get
        {
            if (_range?.GrammaticalErrors != null)
            {
                _grammaticalErrors ??= new WordProofreadingErrors(_range.GrammaticalErrors);
                return _grammaticalErrors;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordSubdocuments? Subdocuments
    {
        get
        {
            if (_range?.Subdocuments != null)
            {
                _subdocuments ??= new WordSubdocuments(_range.Subdocuments);
                return _subdocuments;
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordContentControls? ContentControls
    {
        get
        {
            if (_range?.ContentControls != null)
            {
                _contentControls ??= new WordContentControls(_range.ContentControls);
                return _contentControls;
            }
            return null;
        }
    }

    public IWordConflicts? Conflicts
    {
        get
        {
            if (_range?.Conflicts != null)
            {
                _conflicts ??= new WordConflicts(_range.Conflicts);
                return _conflicts;
            }
            return null;
        }
    }

    public IWordEditors? Editors
    {
        get
        {
            if (_range?.Editors != null)
            {
                _editors ??= new WordEditors(_range.Editors);
                return _editors;
            }
            return null;
        }
    }


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
        get => _range?.EmphasisMark != null ? _range.EmphasisMark.EnumConvert(WdEmphasisMark.wdEmphasisMarkNone) : WdEmphasisMark.wdEmphasisMarkNone;
        set
        {
            if (_range != null) _range.EmphasisMark = value.EnumConvert(MsWord.WdEmphasisMark.wdEmphasisMarkNone);
        }
    }

    /// <inheritdoc/>
    public WdCharacterWidth CharacterWidth
    {
        get => _range?.CharacterWidth != null ? _range.CharacterWidth.EnumConvert(WdCharacterWidth.wdWidthFullWidth) : WdCharacterWidth.wdWidthFullWidth;
        set
        {
            if (_range != null) _range.CharacterWidth = value.EnumConvert(MsWord.WdCharacterWidth.wdWidthFullWidth);
        }
    }

    /// <inheritdoc/>
    public WdHorizontalInVerticalType HorizontalInVertical
    {
        get => _range?.HorizontalInVertical != null ? _range.HorizontalInVertical.EnumConvert(WdHorizontalInVerticalType.wdHorizontalInVerticalNone) : WdHorizontalInVerticalType.wdHorizontalInVerticalNone;
        set
        {
            if (_range != null) _range.HorizontalInVertical = value.EnumConvert(MsWord.WdHorizontalInVerticalType.wdHorizontalInVerticalNone);
        }
    }

    /// <inheritdoc/>
    public WdTextOrientation Orientation
    {
        get => _range?.Orientation != null ? _range.Orientation.EnumConvert(WdTextOrientation.wdTextOrientationHorizontal) : WdTextOrientation.wdTextOrientationHorizontal;
        set
        {
            if (_range != null) _range.Orientation = value.EnumConvert(MsWord.WdTextOrientation.wdTextOrientationHorizontal);
        }
    }

    /// <inheritdoc/>
    public WdTwoLinesInOneType TwoLinesInOne
    {
        get => _range?.TwoLinesInOne != null ? _range.TwoLinesInOne.EnumConvert(WdTwoLinesInOneType.wdTwoLinesInOneNone) : WdTwoLinesInOneType.wdTwoLinesInOneNone;
        set
        {
            if (_range != null) _range.TwoLinesInOne = value.EnumConvert(MsWord.WdTwoLinesInOneType.wdTwoLinesInOneNone);
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageID
    {
        get => _range?.LanguageID != null ? _range.LanguageID.EnumConvert(WdLanguageID.wdLanguageNone) : WdLanguageID.wdLanguageNone;
        set
        {
            if (_range != null) _range.LanguageID = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageIDFarEast
    {
        get => _range?.LanguageIDFarEast != null ? _range.LanguageIDFarEast.EnumConvert(WdLanguageID.wdLanguageNone) : WdLanguageID.wdLanguageNone;
        set
        {
            if (_range != null) _range.LanguageIDFarEast = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
        }
    }

    /// <inheritdoc/>
    public WdLanguageID LanguageIDOther
    {
        get => _range?.LanguageIDOther != null ? _range.LanguageIDOther.EnumConvert(WdLanguageID.wdLanguageNone) : WdLanguageID.wdLanguageNone;
        set
        {
            if (_range != null) _range.LanguageIDOther = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
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
            log.Error($"SaveAsText failed: {ex.Message}", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有延迟初始化的封装对象
            _duplicate?.Dispose();
            _document?.Dispose();
            _font?.Dispose();
            _paragraphFormat?.Dispose();
            _shading?.Dispose();
            _listFormat?.Dispose();
            _pageSetup?.Dispose();
            _paragraphs?.Dispose();
            _sentences?.Dispose();
            _words?.Dispose();
            _characters?.Dispose();
            _tables?.Dispose();
            _bookmarks?.Dispose();
            _fields?.Dispose();
            _hyperlinks?.Dispose();
            _formFields?.Dispose();
            _revisions?.Dispose();
            _comments?.Dispose();
            _footnotes?.Dispose();
            _endnotes?.Dispose();
            _find?.Dispose();
            _formattedText?.Dispose();
            _shapeRangeField?.Dispose();
            _inlineShapes?.Dispose();
            _borders?.Dispose();
            _listParagraphs?.Dispose();
            _readabilityStatistics?.Dispose();
            _spellingErrors?.Dispose();
            _grammaticalErrors?.Dispose();
            _subdocuments?.Dispose();
            _contentControls?.Dispose();
            _conflicts?.Dispose();
            _editors?.Dispose();

            // 释放COM对象
            if (_range != null)
            {
                Marshal.ReleaseComObject(_range);
                _range = null;
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