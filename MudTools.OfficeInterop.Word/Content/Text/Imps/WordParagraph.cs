//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Paragraph 的封装实现类。
/// </summary>
internal class WordParagraph : IWordParagraph
{
    private MsWord.Paragraph _paragraph;

    internal MsWord.Paragraph InternalComObject => _paragraph;

    private bool _disposedValue;

    internal WordParagraph(MsWord.Paragraph paragraph)
    {
        _paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _paragraph != null ? new WordApplication(_paragraph.Application) : null;

    /// <inheritdoc/>
    public object Parent => _paragraph?.Parent;


    /// <inheritdoc/>
    public IWordRange Range => _paragraph?.Range != null ? new WordRange(_paragraph.Range) : null;

    /// <inheritdoc/>
    public WdParagraphAlignment Alignment
    {
        get => _paragraph?.Alignment != null ? (WdParagraphAlignment)(int)_paragraph?.Alignment : WdParagraphAlignment.wdAlignParagraphLeft;
        set
        {
            if (_paragraph != null) _paragraph.Alignment = (MsWord.WdParagraphAlignment)(int)value;
        }
    }

    /// <inheritdoc/>
    public float FirstLineIndent
    {
        get => _paragraph?.FirstLineIndent ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.FirstLineIndent = value;
        }
    }

    /// <inheritdoc/>
    public float LeftIndent
    {
        get => _paragraph?.LeftIndent ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.LeftIndent = value;
        }
    }

    /// <inheritdoc/>
    public float RightIndent
    {
        get => _paragraph?.RightIndent ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.RightIndent = value;
        }
    }

    /// <inheritdoc/>
    public float LineSpacing
    {
        get => _paragraph?.LineSpacing ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.LineSpacing = value;
        }
    }

    /// <inheritdoc/>
    public WdLineSpacing LineSpacingRule
    {
        get => _paragraph?.LineSpacingRule != null ? (WdLineSpacing)(int)_paragraph?.LineSpacingRule : WdLineSpacing.wdLineSpaceSingle;
        set
        {
            if (_paragraph != null) _paragraph.LineSpacingRule = (MsWord.WdLineSpacing)(int)value;
        }
    }

    /// <inheritdoc/>
    public float SpaceBefore
    {
        get => _paragraph?.SpaceBefore ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.SpaceBefore = value;
        }
    }

    /// <inheritdoc/>
    public float SpaceAfter
    {
        get => _paragraph?.SpaceAfter ?? 0f;
        set
        {
            if (_paragraph != null)
                _paragraph.SpaceAfter = value;
        }
    }

    /// <inheritdoc/>
    public int KeepTogether
    {
        get => _paragraph?.KeepTogether ?? 0;
        set
        {
            if (_paragraph != null)
                _paragraph.KeepTogether = value;
        }
    }

    /// <inheritdoc/>
    public int KeepWithNext
    {
        get => _paragraph?.KeepWithNext ?? 0;
        set
        {
            if (_paragraph != null)
                _paragraph.KeepWithNext = value;
        }
    }

    /// <inheritdoc/>
    public int PageBreakBefore
    {
        get => _paragraph?.PageBreakBefore ?? 0;
        set
        {
            if (_paragraph != null)
                _paragraph.PageBreakBefore = value;
        }
    }

    /// <inheritdoc/>
    public WdOutlineLevel OutlineLevel
    {
        get => _paragraph?.OutlineLevel != null ? (WdOutlineLevel)(int)_paragraph?.OutlineLevel : WdOutlineLevel.wdOutlineLevel1;
        set
        {
            if (_paragraph != null) _paragraph.OutlineLevel = (MsWord.WdOutlineLevel)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordTabStops TabStops => _paragraph?.TabStops != null ? new WordTabStops(_paragraph.TabStops) : null;

    /// <inheritdoc/>
    public IWordBorders Borders => _paragraph?.Borders != null ? new WordBorders(_paragraph.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading => _paragraph?.Shading != null ? new WordShading(_paragraph.Shading) : null;

    /// <inheritdoc/>
    public int CharactersCount => _paragraph?.Range?.Characters?.Count ?? 0;

    /// <inheritdoc/>
    public int WordsCount => _paragraph?.Range?.Words?.Count ?? 0;

    /// <inheritdoc/>
    public int SentencesCount => _paragraph?.Range?.Sentences?.Count ?? 0;

    /// <inheritdoc/>
    public bool IsHeading => _paragraph?.OutlineLevel >= MsWord.WdOutlineLevel.wdOutlineLevel1 &&
                            _paragraph?.OutlineLevel <= MsWord.WdOutlineLevel.wdOutlineLevel9;

    /// <inheritdoc/>
    public bool IsEmpty => string.IsNullOrWhiteSpace(GetText());

    /// <inheritdoc/>
    public IWordParagraphFormat Format => _paragraph?.Format != null ? new WordParagraphFormat(_paragraph.Format) : null;

    /// <inheritdoc/>
    public IWordFont Font => _paragraph?.Range?.Font != null ? new WordFont(_paragraph.Range.Font) : null;

    #endregion

    #region 方法实现

    public IWordParagraph Next(int count = 1)
    {
        var paragraph = _paragraph?.Next(count);
        if (paragraph != null)
        {
            return new WordParagraph(paragraph);
        }
        return null;
    }

    public IWordParagraph Previous(int count = 1)
    {
        var paragraph = _paragraph?.Previous(count);
        if (paragraph != null)
        {
            return new WordParagraph(paragraph);
        }
        return null;
    }

    public void OutlinePromote()
    {
        _paragraph?.OutlinePromote();
    }

    public void OutlineDemote()
    {
        _paragraph?.OutlineDemote();
    }

    public void OutlineDemoteToBody()
    {
        _paragraph?.OutlineDemoteToBody();
    }
    public void Indent()
    {
        _paragraph?.Indent();
    }

    public void Outdent()
    {
        _paragraph?.Outdent();
    }

    public void SelectNumber()
    {
        _paragraph?.SelectNumber();
    }
    public void ResetAdvanceTo()
    {
        _paragraph?.ResetAdvanceTo();
    }

    public void ListAdvanceTo(short Level1 = 0, short Level2 = 0, short Level3 = 0, short Level4 = 0, short Level5 = 0, short Level6 = 0, short Level7 = 0, short Level8 = 0, short Level9 = 0)
    {
        _paragraph?.ListAdvanceTo(Level1, Level2, Level3, Level4, Level5, Level6, Level7, Level8, Level9);
    }

    public void SeparateList()
    {
        _paragraph?.SeparateList();
    }

    public void JoinList()
    {
        _paragraph?.JoinList();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _paragraph?.Range?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _paragraph?.Range?.Delete();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _paragraph?.Range?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _paragraph?.Range?.Cut();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _paragraph?.Range?.Paste();
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _paragraph?.Range?.Text?.TrimEnd('\r') ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_paragraph?.Range != null)
        {
            _paragraph.Range.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public void AppendText(string text)
    {
        if (_paragraph?.Range != null && !string.IsNullOrEmpty(text))
        {
            var endRange = _paragraph.Range.Duplicate;
            endRange.Collapse(MsWord.WdCollapseDirection.wdCollapseEnd);
            endRange.Text = text;
        }
    }

    /// <inheritdoc/>
    public void InsertBefore(string text)
    {
        if (_paragraph?.Range != null && !string.IsNullOrEmpty(text))
        {
            var startRange = _paragraph.Range.Duplicate;
            startRange.Collapse(MsWord.WdCollapseDirection.wdCollapseStart);
            startRange.Text = text;
        }
    }

    /// <inheritdoc/>
    public void InsertAfter(string text)
    {
        if (_paragraph?.Range != null && !string.IsNullOrEmpty(text))
        {
            var endRange = _paragraph.Range.Duplicate;
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
            string paragraphText = GetText();
            return matchCase ?
                paragraphText.Contains(text) :
                paragraphText.ToLower().Contains(text.ToLower());
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public int ReplaceText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
    {
        if (_paragraph?.Range == null || string.IsNullOrEmpty(findText)) return 0;

        int replaceCount = 0;
        try
        {
            var findObj = _paragraph.Range.Find;
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
    public void SetIndent(float leftIndent, float firstLineIndent = 0, float rightIndent = 0)
    {
        if (_paragraph != null)
        {
            _paragraph.LeftIndent = leftIndent;
            _paragraph.FirstLineIndent = firstLineIndent;
            _paragraph.RightIndent = rightIndent;
        }
    }

    /// <inheritdoc/>
    public void SetSpacing(float before, float after)
    {
        if (_paragraph != null)
        {
            _paragraph.SpaceBefore = before;
            _paragraph.SpaceAfter = after;
        }
    }

    /// <inheritdoc/>
    public void SetLineSpacing(float lineSpacing, MsWord.WdLineSpacing rule = MsWord.WdLineSpacing.wdLineSpaceSingle)
    {
        if (_paragraph != null)
        {
            _paragraph.LineSpacing = lineSpacing;
            _paragraph.LineSpacingRule = rule;
        }
    }

    /// <inheritdoc/>
    public void AddTabStop(float position, MsWord.WdTabAlignment alignment = MsWord.WdTabAlignment.wdAlignTabLeft,
                          MsWord.WdTabLeader leader = MsWord.WdTabLeader.wdTabLeaderSpaces)
    {
        _paragraph?.TabStops?.Add(position, alignment, leader);
    }

    /// <inheritdoc/>
    public void ClearTabStops()
    {
        _paragraph?.TabStops?.ClearAll();
    }

    /// <inheritdoc/>
    public void ApplyBulletList()
    {
        if (_paragraph?.Range != null)
        {
            var listTemplateObj = _paragraph.Application.ListGalleries[MsWord.WdListGalleryType.wdBulletGallery].ListTemplates[1];
            _paragraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplateObj,
                ContinuePreviousList: false,
                ApplyTo: MsWord.WdListApplyTo.wdListApplyToWholeList,
                DefaultListBehavior: MsWord.WdDefaultListBehavior.wdWord10ListBehavior);
        }
    }

    /// <inheritdoc/>
    public void ApplyNumberedList()
    {
        if (_paragraph?.Range != null)
        {
            var listTemplateObj = _paragraph.Application.ListGalleries[MsWord.WdListGalleryType.wdNumberGallery].ListTemplates[1];
            _paragraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplateObj,
                ContinuePreviousList: false,
                ApplyTo: MsWord.WdListApplyTo.wdListApplyToWholeList,
                DefaultListBehavior: MsWord.WdDefaultListBehavior.wdWord10ListBehavior);
        }
    }


    /// <inheritdoc/>
    public void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth050pt, WdColor color = WdColor.wdColorGray25)
    {
        if (_paragraph?.Borders != null)
        {
            _paragraph.Borders.Enable = 1;
            foreach (MsWord.Border border in _paragraph.Borders)
            {
                border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                border.Color = (MsWord.WdColor)(int)color;
            }
        }
    }

    /// <inheritdoc/>
    public void RemoveBorders()
    {
        _paragraph.Borders.Enable = 0;
    }

    /// <inheritdoc/>
    public void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_paragraph?.Shading != null)
        {
            _paragraph.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
            if (foregroundColor != WdColor.wdColorAutomatic)
                _paragraph.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
            if (backgroundColor != WdColor.wdColorWhite)
                _paragraph.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;
        }
    }

    /// <inheritdoc/>
    public void RemoveShading()
    {
        _paragraph.Shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
    }

    /// <inheritdoc/>
    public List<(float Position, MsWord.WdTabAlignment Alignment, MsWord.WdTabLeader Leader)> GetTabStopsInfo()
    {
        var tabStopsInfo = new List<(float, MsWord.WdTabAlignment, MsWord.WdTabLeader)>();
        if (_paragraph?.TabStops != null)
        {
            try
            {
                foreach (MsWord.TabStop tabStop in _paragraph.TabStops)
                {
                    tabStopsInfo.Add((tabStop.Position, tabStop.Alignment, tabStop.Leader));
                }
            }
            catch
            {
                // 获取制表符信息失败返回空列表
            }
        }
        return tabStopsInfo;
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            (Range as IDisposable)?.Dispose();
            (Format as IDisposable)?.Dispose();
            (Font as IDisposable)?.Dispose();

            if (_paragraph != null)
            {
                Marshal.ReleaseComObject(_paragraph);
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