//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档实现类
/// </summary>
internal class WordDocument : IWordDocument
{
    private readonly MsWord.Document _document;
    private readonly IWordApplication _application;
    private bool _disposedValue;
    private IWordWindow _activeWindow;
    private IWordSelection _selection;
    private IWordRange _range;
    private IWordStoryRanges _storyRanges;
    private IWordBookmarks _bookmarks;
    private IWordTables _tables;
    private IWordParagraphs _paragraphs;
    private IWordSections _sections;
    private IWordStyles _styles;
    private IWordListTemplates _listTemplates;
    private IWordVariables _variables;
    private IWordCustomProperties _customProperties;

    public IWordApplication Application => _application;

    public string Name => _document.Name;

    public string FullName => _document.FullName;

    public string Title
    {
        get
        {
            return GetBuiltInDocumentProperty("Title");
        }
        set
        {
            SetBuiltInDocumentProperty("Title", value);
        }
    }

    private string GetBuiltInDocumentProperty(string propertyName)
    {
        try
        {
            // 使用反射获取内置文档属性
            var properties = _document.BuiltInDocumentProperties;
            var type = properties.GetType();
            var property = type.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, properties, new object[] { propertyName });

            if (property != null)
            {
                var value = property.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, property, null);
                return value?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void SetBuiltInDocumentProperty(string propertyName, string value)
    {
        try
        {
            // 使用反射设置内置文档属性
            var properties = _document.BuiltInDocumentProperties;
            var type = properties.GetType();
            var property = type.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, properties, new object[] { propertyName });

            if (property != null)
            {
                property.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, property, new object[] { value ?? string.Empty });
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to set document property '{propertyName}'.", ex);
        }
    }

    public string Path => _document.Path;

    public bool Saved
    {
        get => _document.Saved;
        set => _document.Saved = value;
    }

    public bool ReadOnly => _document.ReadOnly;

    public WdProtectionType ProtectionType => (WdProtectionType)_document.ProtectionType;

    public int PageCount => (int)_document.Range().Information[MsWord.WdInformation.wdNumberOfPagesInDocument];

    public int WordCount => _document.Words.Count;

    public int ParagraphCount => _document.Paragraphs.Count;

    public int TableCount => _document.Tables.Count;

    public int BookmarkCount => _document.Bookmarks.Count;

    public object Parent => _document.Parent;

    public IWordInlineShapes? InlineShapes => _document != null ? new WordInlineShapes(_document.InlineShapes) : null;

    public IWordShapes? Shapes => _document != null ? new WordShapes(_document.Shapes) : null;

    public IWordWindow ActiveWindow
    {
        get
        {
            _activeWindow ??= new WordWindow(_document.ActiveWindow);
            return _activeWindow;
        }
    }

    public IWordSelection Selection
    {
        get
        {
            if (_selection == null)
            {
                _selection = new WordSelection(_document.Application.Selection, this);
            }
            return _selection;
        }
    }

    public IWordRange Range
    {
        get
        {
            if (_range == null)
            {
                _range = new WordRange(_document.Range());
            }
            return _range;
        }
    }

    public IWordStoryRanges StoryRanges
    {
        get
        {
            if (_storyRanges == null)
            {
                _storyRanges = new WordStoryRanges(_document.StoryRanges);
            }
            return _storyRanges;
        }
    }

    public IWordBookmarks Bookmarks
    {
        get
        {
            if (_bookmarks == null)
            {
                _bookmarks = new WordBookmarks(_document.Bookmarks, this);
            }
            return _bookmarks;
        }
    }

    public IWordTables Tables
    {
        get
        {
            if (_tables == null)
            {
                _tables = new WordTables(_document.Tables, this);
            }
            return _tables;
        }
    }

    public IWordParagraphs Paragraphs
    {
        get
        {
            if (_paragraphs == null)
            {
                _paragraphs = new WordParagraphs(_document.Paragraphs, this);
            }
            return _paragraphs;
        }
    }

    public IWordSections Sections
    {
        get
        {
            if (_sections == null)
            {
                _sections = new WordSections(_document.Sections, this);
            }
            return _sections;
        }
    }

    public IWordStyles Styles
    {
        get
        {
            _styles ??= new WordStyles(_document.Styles);
            return _styles;
        }
    }

    public IWordListTemplates ListTemplates
    {
        get
        {
            _listTemplates ??= new WordListTemplates(_document.ListTemplates, this);
            return _listTemplates;
        }
    }

    public IWordVariables Variables
    {
        get
        {
            _variables ??= new WordVariables(_document.Variables, this);
            return _variables;
        }
    }

    public IWordCustomProperties CustomProperties
    {
        get
        {
            _customProperties ??= new WordCustomProperties(_document.CustomDocumentProperties, this);
            return _customProperties;
        }
    }

    public WdViewType ViewType
    {
        get => (WdViewType)_document.ActiveWindow.View.Type;
        set => _document.ActiveWindow.View.Type = (MsWord.WdViewType)value;
    }

    public bool ShowParagraphs
    {
        get => _document.ActiveWindow.View.ShowParagraphs;
        set => _document.ActiveWindow.View.ShowParagraphs = value;
    }

    public bool ShowHiddenText
    {
        get => _document.ActiveWindow.View.ShowHiddenText;
        set => _document.ActiveWindow.View.ShowHiddenText = value;
    }

    public string Password
    {
        set => _document.Password = value;
    }

    public string WritePassword
    {
        set => _document.WritePassword = value;
    }

    public IWordRange this[int index] => Range;

    public IWordRange this[string bookmarkName] => GetBookmark(bookmarkName)?.Range;

    internal WordDocument(MsWord.Document document, IWordApplication application)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _application = application;
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _document.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate document.", ex);
        }
    }

    public void Save(string fileName = null, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault)
    {
        try
        {
            if (string.IsNullOrEmpty(fileName))
            {
                _document.Save();
            }
            else
            {
                _document.SaveAs2(fileName, (MsWord.WdSaveFormat)fileFormat);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save document.", ex);
        }
    }

    public void SaveAs(string fileName, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault, bool readOnlyRecommended = false)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            var readOnly = readOnlyRecommended ? (object)true : missing;
            _document.SaveAs2(fileName, (MsWord.WdSaveFormat)fileFormat, ref readOnly);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save document as.", ex);
        }
    }

    public void Close(bool saveChanges = true)
    {
        try
        {
            var saveOption = saveChanges ? MsWord.WdSaveOptions.wdSaveChanges : MsWord.WdSaveOptions.wdDoNotSaveChanges;
            _document.Close(saveOption);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close document.", ex);
        }
    }

    public void PrintOut(int copies = 1, string pages = "")
    {
        try
        {
            object background = missing;
            object append = missing;
            object range = missing;
            object outputFileName = missing;
            object from = missing;
            object to = missing;
            object item = missing;
            object copiesObj = copies;
            object pagesObj = string.IsNullOrEmpty(pages) ? missing : (object)pages;
            object pageType = missing;
            object printToFile = missing;
            object collate = missing;
            object fileName = missing;
            object lineNumbers = missing;
            object summaryLen = missing;
            object wordDialog = missing;

            _document.PrintOut(
                ref background,
                ref append,
                ref range,
                ref outputFileName,
                ref from,
                ref to,
                ref item,
                ref copiesObj,
                ref pagesObj,
                ref pageType,
                ref printToFile,
                ref collate,
                ref fileName,
                ref lineNumbers,
                ref summaryLen,
                ref wordDialog);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to print document.", ex);
        }
    }
    public void Protect(WdProtectionType protectionType, string password = null)
    {
        try
        {
            MsWord.WdProtectionType protectionTypeObj = (MsWord.WdProtectionType)protectionType;
            object noResetObj = missing;
            object passwordObj = string.IsNullOrEmpty(password) ? missing : (object)password;

            _document.Protect(
                protectionTypeObj,
                ref noResetObj,
                ref passwordObj);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to protect document.", ex);
        }
    }

    public void Unprotect(string password = null)
    {
        try
        {
            var passwordObj = string.IsNullOrEmpty(password) ? missing : (object)password;
            _document.Unprotect(ref passwordObj);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to unprotect document.", ex);
        }
    }

    public bool IsProtected()
    {
        try
        {
            return _document.ProtectionType != MsWord.WdProtectionType.wdNoProtection;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to check document protection status.", ex);
        }
    }

    public string GetRangeText(int start, int end)
    {
        if (start < 0 || end < start)
            throw new ArgumentOutOfRangeException("Invalid range parameters.");

        try
        {
            var range = _document.Range(start, end);
            return range.Text;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get range text.", ex);
        }
    }

    public void SetRangeText(int start, int end, string text)
    {
        if (start < 0 || end < start)
            throw new ArgumentOutOfRangeException("Invalid range parameters.");

        try
        {
            var range = _document.Range(start, end);
            range.Text = text ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set range text.", ex);
        }
    }

    public void InsertText(int position, string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            range.Text = text;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert text.", ex);
        }
    }

    public void InsertFile(string fileName, int position = -1)
    {
        if (!File.Exists(fileName))
            throw new FileNotFoundException("File not found.", fileName);

        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            range.InsertFile(fileName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to insert file.", ex);
        }
    }

    public int FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        try
        {
            var find = _document.Content.Find;
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

            var count = 0;
            object replaceAll = MsWord.WdReplace.wdReplaceAll;

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
            object replaceObj = replaceAll;
            object matchKashidaObj = missing;
            object matchDiacriticsObj = missing;
            object matchAlefHamzaObj = missing;
            object matchControlObj = missing;

            find.Execute(
                ref findTextObj, ref matchCaseObj, ref matchWholeWordObj, ref matchWildcardsObj,
                ref matchSoundsLikeObj, ref matchAllWordFormsObj, ref forwardObj, ref wrapObj,
                ref formatObj, ref replaceWithObj, ref replaceObj, ref matchKashidaObj,
                ref matchDiacriticsObj, ref matchAlefHamzaObj, ref matchControlObj);

            return 1;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to find and replace text.", ex);
        }
    }

    public IWordBookmark AddBookmark(string name, int start, int end)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            var range = _document.Range(start, end);
            var bookmark = _document.Bookmarks.Add(name, range);
            return new WordBookmark(bookmark);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add bookmark '{name}'.", ex);
        }
    }

    public IWordBookmark GetBookmark(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            var bookmark = _document.Bookmarks[name];
            return bookmark != null ? new WordBookmark(bookmark) : null;
        }
        catch
        {
            return null;
        }
    }

    public void DeleteBookmark(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            if (_document.Bookmarks.Exists(name))
            {
                _document.Bookmarks[name].Delete();
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete bookmark '{name}'.", ex);
        }
    }

    public IWordTable AddTable(int rows, int columns, int position = -1)
    {
        if (rows <= 0 || columns <= 0)
            throw new ArgumentException("Rows and columns must be greater than zero.");

        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            var table = _document.Tables.Add(range, rows, columns);
            return new WordTable(table);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add table.", ex);
        }
    }

    public IWordParagraph AddParagraph(int position, string text = "")
    {
        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            var paragraph = _document.Paragraphs.Add(range);
            if (!string.IsNullOrEmpty(text))
            {
                paragraph.Range.Text = text;
            }
            return new WordParagraph(paragraph);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add paragraph.", ex);
        }
    }

    public void AddSectionBreak(int position, int type = 2)
    {
        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            range.InsertBreak((MsWord.WdBreakType)type);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add section break.", ex);
        }
    }

    public void AddPageBreak(int position)
    {
        try
        {
            var range = position >= 0 ? _document.Range(position, position) : _document.Range();
            range.InsertBreak(MsWord.WdBreakType.wdPageBreak);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add page break.", ex);
        }
    }

    public void AddHeader(string text, bool primary = true)
    {
        try
        {
            var section = _document.Sections[1];
            var header = primary ? section.Headers[MsWord.WdHeaderFooterIndex.wdHeaderFooterPrimary] : section.Headers[MsWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
            header.Range.Text = text;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add header.", ex);
        }
    }

    public void AddFooter(string text, bool primary = true)
    {
        try
        {
            var section = _document.Sections[1];
            var footer = primary ? section.Footers[MsWord.WdHeaderFooterIndex.wdHeaderFooterPrimary] : section.Footers[MsWord.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
            footer.Range.Text = text;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add footer.", ex);
        }
    }

    public void SetMargins(float top, float bottom, float left, float right)
    {
        try
        {
            _document.PageSetup.TopMargin = top;
            _document.PageSetup.BottomMargin = bottom;
            _document.PageSetup.LeftMargin = left;
            _document.PageSetup.RightMargin = right;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set margins.", ex);
        }
    }

    public void SetPageOrientation(bool landscape = false)
    {
        try
        {
            _document.PageSetup.Orientation = landscape ? MsWord.WdOrientation.wdOrientLandscape : MsWord.WdOrientation.wdOrientPortrait;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set page orientation.", ex);
        }
    }

    public void SetPageSize(float width, float height)
    {
        try
        {
            _document.PageSetup.PageWidth = width;
            _document.PageSetup.PageHeight = height;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set page size.", ex);
        }
    }

    public IWordVariable AddVariable(string name, string value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            var variable = _document.Variables.Add(name, value ?? string.Empty);
            return new WordVariable(variable);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add variable '{name}'.", ex);
        }
    }

    public string GetVariable(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            return _document.Variables[name]?.Value;
        }
        catch
        {
            return null;
        }
    }

    public void DeleteVariable(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            _document.Variables[name]?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete variable '{name}'.", ex);
        }
    }

    public void UpdateAllFields()
    {
        try
        {
            _document.Fields.Update();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to update fields.", ex);
        }
    }

    public void AcceptAllRevisions()
    {
        try
        {
            _document.AcceptAllRevisions();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to accept all revisions.", ex);
        }
    }

    public void RejectAllRevisions()
    {
        try
        {
            _document.RejectAllRevisions();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reject all revisions.", ex);
        }
    }


    public void ExportAsPdf(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            _document.ExportAsFixedFormat(fileName, MsWord.WdExportFormat.wdExportFormatPDF);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to export as PDF.", ex);
        }
    }


    public void Refresh()
    {
        try
        {
            _document.Repaginate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh document.", ex);
        }
    }

    public IEnumerator<IWordRange> GetEnumerator()
    {
        return new List<IWordRange> { Range }.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private static readonly object missing = System.Reflection.Missing.Value;
    private static readonly object replaceAll = MsWord.WdReplace.wdReplaceAll;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _activeWindow?.Dispose();
            _selection?.Dispose();
            _range?.Dispose();
            _storyRanges?.Dispose();
            _bookmarks?.Dispose();
            _tables?.Dispose();
            _paragraphs?.Dispose();
            _sections?.Dispose();
            _styles?.Dispose();
            _listTemplates?.Dispose();
            _variables?.Dispose();
            _customProperties?.Dispose();
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
