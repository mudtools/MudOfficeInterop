//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word Selection 实现类
/// </summary>
internal class WordSelection : IWordSelection
{
    private readonly MsWord.Selection _selection;
    private bool _disposedValue;
    private IWordFind _find;
    private IWordRange _range;

    public IWordApplication? Application => _selection != null ? new WordApplication(_selection.Application) : null;


    public string Text
    {
        get => _selection.Text;
        set => _selection.Text = value;
    }

    public WdBuiltinStyle Style
    {
        get => (WdBuiltinStyle)(int)_selection.get_Style();
        set => _selection.set_Style(value.EnumConvert(MsWord.WdBuiltinStyle.wdStyleBodyText));
    }

    public WdTextOrientation Orientation
    {
        get => _selection.Orientation.EnumConvert(WdTextOrientation.wdTextOrientationHorizontal);
        set => _selection.Orientation = value.EnumConvert(MsWord.WdTextOrientation.wdTextOrientationHorizontal);
    }

    public WdSelectionType Type => _selection.Type.EnumConvert(WdSelectionType.wdNoSelection);

    public WdStoryType StoryType => _selection.StoryType.EnumConvert(WdStoryType.wdMainTextStory);

    public WdSelectionFlags Flags => _selection.Flags.EnumConvert(WdSelectionFlags.wdSelReplace);

    public WdLanguageID LanguageID
    {
        get => _selection.LanguageID.EnumConvert(WdLanguageID.wdLanguageNone);
        set => _selection.LanguageID = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
    }

    public WdLanguageID LanguageIDFarEast
    {
        get => _selection.LanguageIDFarEast.EnumConvert(WdLanguageID.wdLanguageNone);
        set => _selection.LanguageIDFarEast = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
    }

    public WdLanguageID LanguageIDOther
    {
        get => _selection.LanguageIDOther.EnumConvert(WdLanguageID.wdLanguageNone);
        set => _selection.LanguageIDOther = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
    }
    public int Start
    {
        get => _selection.Start;
        set => _selection.Start = value;
    }

    public int End
    {
        get => _selection.End;
        set => _selection.End = value;
    }

    public int StoryLength
    {
        get => _selection.StoryLength;
    }

    public int BookmarkID
    {
        get => _selection.BookmarkID;
    }

    public int PreviousBookmarkID
    {
        get => _selection.PreviousBookmarkID;
    }

    public bool IsEndOfRowMark
    {
        get => _selection.IsEndOfRowMark;
    }

    public bool Active
    {
        get => _selection.Active;
    }

    public bool StartIsActive
    {
        get => _selection.StartIsActive;
        set => _selection.StartIsActive = value;
    }

    public bool ExtendMode
    {
        get => _selection.ExtendMode;
        set => _selection.ExtendMode = value;
    }

    public bool ColumnSelectMode
    {
        get => _selection.ColumnSelectMode;
        set => _selection.ColumnSelectMode = value;
    }

    public bool IPAtEndOfLine
    {
        get => _selection.IPAtEndOfLine;
    }

    public int Length => _selection.End - _selection.Start;

    public object Parent => _selection.Parent;

    public IWordDocument? Document => _selection != null ? new WordDocument(_selection.Document) : null;

    public IWordFont? Font => _selection != null ? new WordFont(_selection.Font) : null;

    public IWordShapeRange? ShapeRange => _selection != null ? new WordShapeRange(_selection.ShapeRange) : null;

    public IWordInlineShapes? InlineShapes => _selection != null ? new WordInlineShapes(_selection.InlineShapes) : null;

    public IWordParagraphs? Paragraphs => _selection != null ? new WordParagraphs(_selection.Paragraphs) : null;

    public IWordBorders? Borders => _selection != null ? new WordBorders(_selection.Borders) : null;

    public IWordShading? Shading => _selection != null ? new WordShading(_selection.Shading) : null;

    public IWordFields? Fields => _selection != null ? new WordFields(_selection.Fields) : null;

    public IWordFormFields? FormFields => _selection != null ? new WordFormFields(_selection.FormFields) : null;

    public IWordFrames? Frames => _selection != null ? new WordFrames(_selection.Frames) : null;

    public IWordParagraphFormat? ParagraphFormat => _selection != null ? new WordParagraphFormat(_selection.ParagraphFormat) : null;

    public IWordPageSetup? PageSetup => _selection != null ? new WordPageSetup(_selection.PageSetup) : null;

    public IWordBookmarks Bookmarks => _selection != null ? new WordBookmarks(_selection.Bookmarks) : null;

    public IWordSections Sections => _selection != null ? new WordSections(_selection.Sections) : null;

    public IWordCells Cells => _selection != null ? new WordCells(_selection.Cells) : null;

    public IWordColumns Columns => _selection != null ? new WordColumns(_selection.Columns) : null;

    public IWordRows Rows => _selection != null ? new WordRows(_selection.Rows) : null;

    public IWordHeaderFooter HeaderFooter => _selection != null ? new WordHeaderFooter(_selection.HeaderFooter) : null;

    public IWordComments Comments => _selection != null ? new WordComments(_selection.Comments) : null;

    public IWordEndnotes Endnotes => _selection != null ? new WordEndnotes(_selection.Endnotes) : null;

    public IWordFootnotes Footnotes => _selection != null ? new WordFootnotes(_selection.Footnotes) : null;

    public IWordCharacters Characters => _selection != null ? new WordCharacters(_selection.Characters) : null;

    public IWordSentences Sentences => _selection != null ? new WordSentences(_selection.Sentences) : null;

    public IWordWords Words => _selection != null ? new WordWords(_selection.Words) : null;

    public IWordTables Tables => _selection != null ? new WordTables(_selection.Tables) : null;

    public IWordRange FormattedText => _selection != null ? new WordRange(_selection.FormattedText) : null;


    public string FontName
    {
        get => _selection.Font.Name;
        set => _selection.Font.Name = value;
    }

    public float FontSize
    {
        get => _selection.Font.Size;
        set => _selection.Font.Size = value;
    }

    public bool Bold
    {
        get => _selection.Font.Bold == 1;
        set => _selection.Font.Bold = value ? 1 : 0;
    }

    public bool Italic
    {
        get => _selection.Font.Italic == 1;
        set => _selection.Font.Italic = value ? 1 : 0;
    }

    public int Underline
    {
        get => (int)_selection.Font.Underline;
        set => _selection.Font.Underline = (MsWord.WdUnderline)value;
    }

    public WdColor FontColor
    {
        get => (WdColor)_selection.Font.Color;
        set => _selection.Font.Color = (MsWord.WdColor)value;
    }

    public int Alignment
    {
        get => (int)_selection.ParagraphFormat.Alignment;
        set => _selection.ParagraphFormat.Alignment = (MsWord.WdParagraphAlignment)value;
    }

    public float LineSpacing
    {
        get => _selection.ParagraphFormat.LineSpacing;
        set => _selection.ParagraphFormat.LineSpacing = value;
    }

    public float SpaceBefore
    {
        get => _selection.ParagraphFormat.SpaceBefore;
        set => _selection.ParagraphFormat.SpaceBefore = value;
    }

    public float SpaceAfter
    {
        get => _selection.ParagraphFormat.SpaceAfter;
        set => _selection.ParagraphFormat.SpaceAfter = value;
    }

    public float FirstLineIndent
    {
        get => _selection.ParagraphFormat.FirstLineIndent;
        set => _selection.ParagraphFormat.FirstLineIndent = value;
    }

    public IWordFind Find
    {
        get
        {
            if (_find == null)
            {
                _find = new WordFind(_selection.Find);
            }
            return _find;
        }
    }

    public IWordRange Range
    {
        get
        {
            if (_range == null)
            {
                _range = new WordRange(_selection.Range);
            }
            return _range;
        }
    }

    internal WordSelection(MsWord.Selection selection)
    {
        _selection = selection ?? throw new ArgumentNullException(nameof(selection));
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _selection.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate selection.", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _selection.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy selection.", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _selection.Cut();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to cut selection.", ex);
        }
    }

    public void Paste()
    {
        try
        {
            _selection.Paste();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to paste content.", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _selection.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete selection.", ex);
        }
    }

    public void ClearFormatting()
    {
        try
        {
            _selection.ClearFormatting();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to clear formatting.", ex);
        }
    }

    public void InsertText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        try
        {
            _selection.TypeText(text);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert text.", ex);
        }
    }

    public void InsertParagraph()
    {
        try
        {
            _selection.InsertParagraph();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert paragraph.", ex);
        }
    }

    public void InsertParagraphAfter()
    {
        try
        {
            _selection.InsertParagraphAfter();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert paragraph.", ex);
        }
    }

    public void InsertParagraphBefore()
    {
        try
        {
            _selection.InsertParagraphBefore();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert paragraph.", ex);
        }
    }

    public void InsertBefore(string text)
    {
        _selection?.InsertBefore(text);
    }
    public void InsertAfter(string text)
    {
        _selection?.InsertAfter(text);
    }


    public void InsertLineBreak()
    {
        try
        {
            _selection.TypeText("\n");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert line break.", ex);
        }
    }

    public void InsertNewPage()
    {
        _selection?.InsertNewPage();
    }

    public void InsertPageBreak()
    {
        try
        {
            _selection?.InsertBreak(MsWord.WdBreakType.wdPageBreak);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert page break.", ex);
        }
    }

    public IWordRange? Next(WdUnits unit, int count)
    {
        var range = _selection?.Next(unit.EnumConvert(MsWord.WdUnits.wdLine), count);
        return range != null ? new WordRange(range) : null;
    }

    public IWordRange? Previous(WdUnits unit, int count)
    {
        var range = _selection?.Previous(unit.EnumConvert(MsWord.WdUnits.wdLine), count);
        return range != null ? new WordRange(range) : null;
    }

    public IWordTable InsertTable(int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
            throw new ArgumentException("Rows and columns must be greater than zero.");

        try
        {
            var table = _selection.Tables.Add(_selection.Range, rows, columns);
            return new WordTable(table);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert table.", ex);
        }
    }

    public void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool RTF)
    {
        _selection?.PasteExcelTable(linkedToExcel, wordFormatting, RTF);
    }

    public void PasteFormat()
    {
        _selection?.PasteFormat();
    }

    public void PasteAsNestedTable()
    {
        _selection?.PasteAsNestedTable();
    }

    public void PasteSpecial(bool? iconIndex,
        bool? link, WdOLEPlacement? placement,
        bool? displayAsIcon, WdPasteDataType? dataType,
        bool? iconFileName, bool? iconLabel)
    {
        _selection?.PasteSpecial(
            IconIndex: iconIndex.ComArgsVal(),
            Link: link.ComArgsVal(),
            Placement: placement.ComArgsConvert(d => d.EnumConvert(MsWord.WdOLEPlacement.wdInLine)),
            DisplayAsIcon: displayAsIcon.ComArgsVal(),
            DataType: dataType.ComArgsConvert(d => d.EnumConvert(MsWord.WdPasteDataType.wdPasteText)),
            IconFileName: iconFileName.ComArgsVal(),
            IconLabel: iconLabel.ComArgsVal()
            );
    }

    public void PasteAppendTable()
    {
        _selection?.PasteAppendTable();
    }

    public void PasteAndFormat(WdRecoveryType Type)
    {
        _selection?.PasteAndFormat(Type.EnumConvert(MsWord.WdRecoveryType.wdFormatPlainText));
    }

    public int MoveLeft(int unit = 1, int count = 1)
    {
        try
        {
            return _selection.MoveLeft((MsWord.WdUnits)unit, count);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to move selection left.", ex);
        }
    }

    public int MoveRight(int unit = 1, int count = 1)
    {
        try
        {
            return _selection.MoveRight((MsWord.WdUnits)unit, count);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to move selection right.", ex);
        }
    }

    public int MoveUp(int unit = 1, int count = 1)
    {
        try
        {
            return _selection.MoveUp((MsWord.WdUnits)unit, count);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to move selection up.", ex);
        }
    }

    public int MoveDown(int unit = 1, int count = 1)
    {
        try
        {
            return _selection.MoveDown((MsWord.WdUnits)unit, count);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to move selection down.", ex);
        }
    }

    public void SetRange(int start, int end)
    {
        _selection.SetRange(start, end);
    }

    public void Select()
    {
        try
        {
            _selection.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select  content.", ex);
        }
    }
    public bool InRange(IWordRange range)
    {
        if (_selection == null || range == null)
            return false;

        return _selection.InRange(((WordRange)range)._range);
    }

    public void Shrink()
    {
        _selection?.Shrink();
    }

    public void SplitTable()
    {
        _selection?.SplitTable();
    }

    public int? StartOf(WdUnits? unit, WdMovementType? extend)
    {
        return _selection?.StartOf(
            unit.ComArgsConvert(d => d.EnumConvert(MsWord.WdUnits.wdCharacter)),
            extend.ComArgsConvert(d => d.EnumConvert(MsWord.WdMovementType.wdMove)));
    }

    public void SelectCell()
    {
        _selection?.SelectCell();
    }

    public void SelectColumn()
    {
        _selection?.SelectColumn();
    }

    public void SelectCurrentAlignment()
    {
        _selection?.SelectCurrentAlignment();
    }

    public void SelectCurrentColor()
    {
        _selection?.SelectCurrentColor();
    }

    public void SelectCurrentFont()
    {
        _selection?.SelectCurrentFont();
    }

    public void SelectCurrentIndent()
    {
        _selection?.SelectCurrentIndent();
    }

    public void SelectCurrentSpacing()
    {
        _selection?.SelectCurrentSpacing();
    }

    public void SelectCurrentTabs()
    {
        _selection?.SelectCurrentTabs();
    }

    public void SelectRow()
    {
        _selection?.SelectRow();
    }

    public void SelectAll()
    {
        try
        {
            _selection?.WholeStory();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select all content.", ex);
        }
    }

    public void Collapse()
    {
        try
        {
            _selection.Collapse();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to collapse selection.", ex);
        }
    }

    public void Extend(int unit = 1, int count = 1)
    {
        try
        {
            _selection.MoveEnd((MsWord.WdUnits)unit, count);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to extend selection.", ex);
        }
    }

    public bool FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (string.IsNullOrEmpty(findText))
            return false;

        try
        {
            var find = _selection.Find;
            find.ClearFormatting();
            find.Text = findText;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = replaceText ?? string.Empty;
            find.Forward = true;
            find.Wrap = MsWord.WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = matchCase;
            find.MatchWholeWord = matchWholeWord;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            object replaceAllObj = MsWord.WdReplace.wdReplaceAll;

            // 为所有参数创建本地变量
            object findTextObj = missing;
            object matchCaseObj = missing;
            object matchWholeWordObj = missing;
            object matchWildcardsObj = missing;
            object matchSoundsLikeObj = missing;
            object matchAllWordFormsObj = missing;
            object forwardObj = missing;
            object wrapObj = missing;
            object formatObj = missing;
            object replaceWithObj = missing;
            object replaceObj = replaceAllObj;
            object matchKashidaObj = missing;
            object matchDiacriticsObj = missing;
            object matchAlefHamzaObj = missing;
            object matchControlObj = missing;

            return find.Execute(
                ref findTextObj, ref matchCaseObj, ref matchWholeWordObj, ref matchWildcardsObj,
                ref matchSoundsLikeObj, ref matchAllWordFormsObj, ref forwardObj, ref wrapObj,
                ref formatObj, ref replaceWithObj, ref replaceObj, ref matchKashidaObj,
                ref matchDiacriticsObj, ref matchAlefHamzaObj, ref matchControlObj);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find and replace text.", ex);
        }
    }

    public void SetFont(string fontName = null, float fontSize = 0, bool bold = false, bool italic = false, int underline = 0, int color = 0)
    {
        try
        {
            if (!string.IsNullOrEmpty(fontName))
                _selection.Font.Name = fontName;
            if (fontSize > 0)
                _selection.Font.Size = fontSize;
            _selection.Font.Bold = bold ? 1 : 0;
            _selection.Font.Italic = italic ? 1 : 0;
            if (underline >= 0)
                _selection.Font.Underline = (MsWord.WdUnderline)underline;
            if (color >= 0)
                _selection.Font.Color = (MsWord.WdColor)color;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set font formatting.", ex);
        }
    }

    public void SetParagraph(int alignment = 0, float lineSpacing = 0, float spaceBefore = 0, float spaceAfter = 0, float firstLineIndent = 0)
    {
        try
        {
            if (alignment >= 0)
                _selection.ParagraphFormat.Alignment = (MsWord.WdParagraphAlignment)alignment;
            if (lineSpacing > 0)
                _selection.ParagraphFormat.LineSpacing = lineSpacing;
            if (spaceBefore >= 0)
                _selection.ParagraphFormat.SpaceBefore = spaceBefore;
            if (spaceAfter >= 0)
                _selection.ParagraphFormat.SpaceAfter = spaceAfter;
            if (firstLineIndent != 0)
                _selection.ParagraphFormat.FirstLineIndent = firstLineIndent;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set paragraph formatting.", ex);
        }
    }

    public IWordBookmark GetBookmark(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            var bookmark = _selection.Bookmarks[name];
            return bookmark != null ? new WordBookmark(bookmark) : null;
        }
        catch
        {
            return null;
        }
    }

    public IWordBookmark AddBookmark(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            object range = missing;
            var bookmark = _selection.Bookmarks.Add(name, ref range);
            return new WordBookmark(bookmark);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add bookmark '{name}'.", ex);
        }
    }

    public IWordHyperlink AddHyperlink(string address)
    {
        if (string.IsNullOrEmpty(address))
            throw new ArgumentException("Hyperlink address cannot be null or empty.", nameof(address));

        try
        {
            var hyperlink = _selection.Hyperlinks.Add(_selection.Range, address);
            return new WordHyperlink(hyperlink);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add hyperlink to '{address}'.", ex);
        }
    }

    public void Refresh()
    {
        try
        {
            // Word 中没有直接的刷新方法，这里模拟刷新
            var currentRange = _selection.Range;
            currentRange.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh selection.", ex);
        }
    }

    private static readonly object missing = System.Reflection.Missing.Value;
    private static readonly object replaceAll = MsWord.WdReplace.wdReplaceAll;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _find?.Dispose();
            _range?.Dispose();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}