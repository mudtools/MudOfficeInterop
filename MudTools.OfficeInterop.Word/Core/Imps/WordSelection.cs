//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    private readonly IWordDocument _document;
    private bool _disposedValue;
    private IWordFind _find;
    private IWordRange _range;

    public IWordApplication? Application => _selection != null ? new WordApplication(_selection.Application) : null;


    public string Text
    {
        get => _selection.Text;
        set => _selection.TypeText(value);
    }

    public WdSelectionType Type => (WdSelectionType)_selection.Type;

    public int Start
    {
        get => _selection.Start;
        set => _selection.SetRange(value, _selection.End);
    }

    public int End
    {
        get => _selection.End;
        set => _selection.SetRange(_selection.Start, value);
    }

    public int Length => _selection.End - _selection.Start;

    public object Parent => _selection.Parent;

    public IWordDocument Document => _document;

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

    internal WordSelection(MsWord.Selection selection, IWordDocument document)
    {
        _selection = selection ?? throw new ArgumentNullException(nameof(selection));
        _document = document;
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
            _selection.TypeParagraph();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert paragraph.", ex);
        }
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

    public void InsertPageBreak()
    {
        try
        {
            _selection.InsertBreak(MsWord.WdBreakType.wdPageBreak);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert page break.", ex);
        }
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

    public void SelectAll()
    {
        try
        {
            _selection.WholeStory();
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

    public IEnumerable<IWordBookmark> Bookmarks
    {
        get
        {
            var bookmarks = new List<IWordBookmark>();
            try
            {
                foreach (MsWord.Bookmark bookmark in _selection.Bookmarks)
                {
                    bookmarks.Add(new WordBookmark(bookmark));
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to enumerate bookmarks.", ex);
            }
            return bookmarks;
        }
    }

    public IEnumerable<IWordTable> Tables
    {
        get
        {
            var tables = new List<IWordTable>();
            try
            {
                foreach (MsWord.Table table in _selection.Tables)
                {
                    tables.Add(new WordTable(table));
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to enumerate tables.", ex);
            }
            return tables;
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