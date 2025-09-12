//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档页面设置实现类
/// </summary>
internal class WordPageSetup : IWordPageSetup
{
    internal readonly MsWord.PageSetup _pageSetup;
    private bool _disposedValue;

    #region 页面尺寸和边距设置

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    public float TopMargin
    {
        get => _pageSetup.TopMargin;
        set => _pageSetup.TopMargin = value;
    }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    public float BottomMargin
    {
        get => _pageSetup.BottomMargin;
        set => _pageSetup.BottomMargin = value;
    }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    public float LeftMargin
    {
        get => _pageSetup.LeftMargin;
        set => _pageSetup.LeftMargin = value;
    }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    public float RightMargin
    {
        get => _pageSetup.RightMargin;
        set => _pageSetup.RightMargin = value;
    }

    /// <summary>
    /// 获取或设置页面宽度
    /// </summary>
    public float PageWidth
    {
        get => _pageSetup.PageWidth;
        set => _pageSetup.PageWidth = value;
    }

    /// <summary>
    /// 获取或设置页面高度
    /// </summary>
    public float PageHeight
    {
        get => _pageSetup.PageHeight;
        set => _pageSetup.PageHeight = value;
    }

    #endregion

    #region 页眉页脚设置

    public float HeaderDistance
    {
        get => _pageSetup.HeaderDistance;
        set => _pageSetup.HeaderDistance = value;
    }

    public float FooterDistance
    {
        get => _pageSetup.FooterDistance;
        set => _pageSetup.FooterDistance = value;
    }

    public int OddAndEvenPagesHeaderFooter
    {
        get => _pageSetup.OddAndEvenPagesHeaderFooter;
        set => _pageSetup.OddAndEvenPagesHeaderFooter = value;
    }

    public int DifferentFirstPageHeaderFooter
    {
        get => _pageSetup.DifferentFirstPageHeaderFooter;
        set => _pageSetup.DifferentFirstPageHeaderFooter = value;
    }

    #endregion

    #region 页面方向和布局设置

    public WdOrientation Orientation
    {
        get => (WdOrientation)_pageSetup.Orientation;
        set => _pageSetup.Orientation = (MsWord.WdOrientation)value;
    }

    public WdVerticalAlignment VerticalAlignment
    {
        get => (WdVerticalAlignment)_pageSetup.VerticalAlignment;
        set => _pageSetup.VerticalAlignment = (MsWord.WdVerticalAlignment)value;
    }

    public WdLayoutMode LayoutMode
    {
        get => (WdLayoutMode)_pageSetup.LayoutMode;
        set => _pageSetup.LayoutMode = (MsWord.WdLayoutMode)value;
    }

    #endregion

    #region 装订线设置

    public float Gutter
    {
        get => _pageSetup.Gutter;
        set => _pageSetup.Gutter = value;
    }

    public bool GutterOnTop
    {
        get => _pageSetup.GutterOnTop;
        set => _pageSetup.GutterOnTop = value;
    }

    public WdGutterStyle GutterPos
    {
        get => (WdGutterStyle)_pageSetup.GutterPos;
        set => _pageSetup.GutterPos = (MsWord.WdGutterStyle)value;
    }

    public WdGutterStyleOld GutterStyle
    {
        get => (WdGutterStyleOld)_pageSetup.GutterStyle;
        set => _pageSetup.GutterStyle = (MsWord.WdGutterStyleOld)value;
    }

    #endregion

    #region 文本列和行号设置

    public IWordLineNumbering LineNumbering
    {
        get => new WordLineNumbering(_pageSetup.LineNumbering);
        set => _pageSetup.LineNumbering = ((WordLineNumbering)value)._lineNumbering;
    }

    public IWordTextColumns TextColumns
    {
        get => new WordTextColumns(_pageSetup.TextColumns);
        set => _pageSetup.TextColumns = ((WordTextColumns)value)._textColumns;
    }

    public float CharsLine
    {
        get => _pageSetup.CharsLine;
        set => _pageSetup.CharsLine = value;
    }

    public float LinesPage
    {
        get => _pageSetup.LinesPage;
        set => _pageSetup.LinesPage = value;
    }

    #endregion

    #region 章节和分节设置

    public WdSectionDirection SectionDirection
    {
        get => (WdSectionDirection)_pageSetup.SectionDirection;
        set => _pageSetup.SectionDirection = (MsWord.WdSectionDirection)value;
    }

    public WdSectionStart SectionStart
    {
        get => (WdSectionStart)_pageSetup.SectionStart;
        set => _pageSetup.SectionStart = (MsWord.WdSectionStart)value;
    }

    #endregion

    #region 书籍折页打印设置

    public bool BookFoldPrinting
    {
        get => _pageSetup.BookFoldPrinting;
        set => _pageSetup.BookFoldPrinting = value;
    }

    public int BookFoldPrintingSheets
    {
        get => _pageSetup.BookFoldPrintingSheets;
        set => _pageSetup.BookFoldPrintingSheets = value;
    }

    public bool BookFoldRevPrinting
    {
        get => _pageSetup.BookFoldRevPrinting;
        set => _pageSetup.BookFoldRevPrinting = value;
    }

    #endregion

    #region 纸张和打印设置

    public WdPaperTray FirstPageTray
    {
        get => (WdPaperTray)_pageSetup.FirstPageTray;
        set => _pageSetup.FirstPageTray = (MsWord.WdPaperTray)value;
    }

    public WdPaperTray OtherPagesTray
    {
        get => (WdPaperTray)_pageSetup.OtherPagesTray;
        set => _pageSetup.OtherPagesTray = (MsWord.WdPaperTray)value;
    }

    public WdPaperSize PaperSize
    {
        get => (WdPaperSize)_pageSetup.PaperSize;
        set => _pageSetup.PaperSize = (MsWord.WdPaperSize)value;
    }

    public bool TwoPagesOnOne
    {
        get => _pageSetup.TwoPagesOnOne;
        set => _pageSetup.TwoPagesOnOne = value;
    }

    #endregion

    #region 其他设置

    public int SuppressEndnotes
    {
        get => _pageSetup.SuppressEndnotes;
        set => _pageSetup.SuppressEndnotes = value;
    }

    public bool ShowGrid
    {
        get => _pageSetup.ShowGrid;
        set => _pageSetup.ShowGrid = value;
    }

    #endregion

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="pageSetup">COM PageSetup 对象</param>
    internal WordPageSetup(MsWord.PageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        if (disposing && _pageSetup != null)
        {
            Marshal.ReleaseComObject(_pageSetup);
        }
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}