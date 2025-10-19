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
    internal MsWord.PageSetup? _pageSetup;
    private bool _disposedValue;

    #region 页面尺寸和边距设置

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    public float TopMargin
    {
        get => _pageSetup?.TopMargin ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.TopMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    public float BottomMargin
    {
        get => _pageSetup?.BottomMargin ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.BottomMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    public float LeftMargin
    {
        get => _pageSetup?.LeftMargin ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.LeftMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    public float RightMargin
    {
        get => _pageSetup?.RightMargin ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.RightMargin = value;
        }
    }

    /// <summary>
    /// 获取或设置页面宽度
    /// </summary>
    public float PageWidth
    {
        get => _pageSetup?.PageWidth ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.PageWidth = value;
        }
    }

    /// <summary>
    /// 获取或设置页面高度
    /// </summary>
    public float PageHeight
    {
        get => _pageSetup?.PageHeight ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.PageHeight = value;
        }
    }

    #endregion

    #region 页眉页脚设置

    public float HeaderDistance
    {
        get => _pageSetup?.HeaderDistance ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.HeaderDistance = value;
        }
    }

    public float FooterDistance
    {
        get => _pageSetup?.FooterDistance ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.FooterDistance = value;
        }
    }

    public int OddAndEvenPagesHeaderFooter
    {
        get => _pageSetup?.OddAndEvenPagesHeaderFooter ?? 0;
        set
        {
            if (_pageSetup != null) _pageSetup.OddAndEvenPagesHeaderFooter = value;
        }
    }

    public int DifferentFirstPageHeaderFooter
    {
        get => _pageSetup?.DifferentFirstPageHeaderFooter ?? 0;
        set
        {
            if (_pageSetup != null) _pageSetup.DifferentFirstPageHeaderFooter = value;
        }
    }
    #endregion

    #region 页面方向和布局设置

    public WdPaperSize PageSize
    {
        get => _pageSetup?.PaperSize.EnumConvert(WdPaperSize.wdPaperA4) ?? WdPaperSize.wdPaperA4;
        set
        {
            if (_pageSetup != null) _pageSetup.PaperSize = value.EnumConvert(MsWord.WdPaperSize.wdPaperA4);
        }
    }

    public WdOrientation Orientation
    {
        get => _pageSetup?.Orientation.EnumConvert(WdOrientation.wdOrientPortrait) ?? WdOrientation.wdOrientPortrait;
        set
        {
            if (_pageSetup != null) _pageSetup.Orientation = value.EnumConvert(MsWord.WdOrientation.wdOrientPortrait);
        }
    }

    public WdVerticalAlignment VerticalAlignment
    {
        get => _pageSetup?.VerticalAlignment.EnumConvert(WdVerticalAlignment.wdAlignVerticalTop) ?? WdVerticalAlignment.wdAlignVerticalTop;
        set
        {
            if (_pageSetup != null) _pageSetup.VerticalAlignment = value.EnumConvert(MsWord.WdVerticalAlignment.wdAlignVerticalTop);
        }
    }

    public WdLayoutMode LayoutMode
    {
        get => _pageSetup?.LayoutMode.EnumConvert(WdLayoutMode.wdLayoutModeDefault) ?? WdLayoutMode.wdLayoutModeDefault;
        set
        {
            if (_pageSetup != null) _pageSetup.LayoutMode = value.EnumConvert(MsWord.WdLayoutMode.wdLayoutModeDefault);
        }
    }

    #endregion

    #region 装订线设置

    public float Gutter
    {
        get => _pageSetup?.Gutter ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.Gutter = value;
        }
    }

    public bool GutterOnTop
    {
        get => _pageSetup?.GutterOnTop ?? false;
        set
        {
            if (_pageSetup != null) _pageSetup.GutterOnTop = value;
        }
    }

    public WdGutterStyle GutterPos
    {
        get => _pageSetup?.GutterPos.EnumConvert(WdGutterStyle.wdGutterPosLeft) ?? WdGutterStyle.wdGutterPosLeft;
        set
        {
            if (_pageSetup != null) _pageSetup.GutterPos = value.EnumConvert(MsWord.WdGutterStyle.wdGutterPosLeft);
        }
    }

    public WdGutterStyleOld GutterStyle
    {
        get => _pageSetup?.GutterStyle.EnumConvert(WdGutterStyleOld.wdGutterStyleBidi) ?? WdGutterStyleOld.wdGutterStyleBidi;
        set
        {
            if (_pageSetup != null) _pageSetup.GutterStyle = value.EnumConvert(MsWord.WdGutterStyleOld.wdGutterStyleBidi);
        }
    }

    #endregion

    #region 文本列和行号设置

    public IWordLineNumbering? LineNumbering
    {
        get => _pageSetup != null ? new WordLineNumbering(_pageSetup.LineNumbering) : null;
        set
        {
            if (_pageSetup != null && value != null) _pageSetup.LineNumbering = ((WordLineNumbering)value)._lineNumbering;
        }
    }

    public IWordTextColumns? TextColumns
    {
        get => _pageSetup != null ? new WordTextColumns(_pageSetup.TextColumns) : null;
        set
        {
            if (_pageSetup != null && value != null) _pageSetup.TextColumns = ((WordTextColumns)value)._textColumns;
        }
    }

    public float CharsLine
    {
        get => _pageSetup?.CharsLine ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.CharsLine = value;
        }
    }

    public float LinesPage
    {
        get => _pageSetup?.LinesPage ?? 0f;
        set
        {
            if (_pageSetup != null) _pageSetup.LinesPage = value;
        }
    }

    #endregion

    #region 章节和分节设置

    public WdSectionDirection SectionDirection
    {
        get => _pageSetup?.SectionDirection.EnumConvert(WdSectionDirection.wdSectionDirectionLtr) ?? WdSectionDirection.wdSectionDirectionLtr;
        set
        {
            if (_pageSetup != null) _pageSetup.SectionDirection = value.EnumConvert(MsWord.WdSectionDirection.wdSectionDirectionLtr);
        }
    }

    public WdSectionStart SectionStart
    {
        get => _pageSetup?.SectionStart.EnumConvert(WdSectionStart.wdSectionContinuous) ?? WdSectionStart.wdSectionContinuous;
        set
        {
            if (_pageSetup != null) _pageSetup.SectionStart = value.EnumConvert(MsWord.WdSectionStart.wdSectionContinuous);
        }
    }

    #endregion

    #region 书籍折页打印设置

    public bool BookFoldPrinting
    {
        get => _pageSetup?.BookFoldPrinting ?? false;
        set
        {
            if (_pageSetup != null) _pageSetup.BookFoldPrinting = value;
        }
    }

    public int BookFoldPrintingSheets
    {
        get => _pageSetup?.BookFoldPrintingSheets ?? 0;
        set
        {
            if (_pageSetup != null) _pageSetup.BookFoldPrintingSheets = value;
        }
    }

    public bool BookFoldRevPrinting
    {
        get => _pageSetup?.BookFoldRevPrinting ?? false;
        set
        {
            if (_pageSetup != null) _pageSetup.BookFoldRevPrinting = value;
        }
    }

    #endregion

    #region 纸张和打印设置

    public WdPaperTray FirstPageTray
    {
        get => _pageSetup?.FirstPageTray.EnumConvert(WdPaperTray.wdPrinterDefaultBin) ?? WdPaperTray.wdPrinterDefaultBin;
        set
        {
            if (_pageSetup != null) _pageSetup.FirstPageTray = value.EnumConvert(MsWord.WdPaperTray.wdPrinterDefaultBin);
        }
    }

    public WdPaperTray OtherPagesTray
    {
        get => _pageSetup?.OtherPagesTray.EnumConvert(WdPaperTray.wdPrinterDefaultBin) ?? WdPaperTray.wdPrinterDefaultBin;
        set
        {
            if (_pageSetup != null) _pageSetup.OtherPagesTray = value.EnumConvert(MsWord.WdPaperTray.wdPrinterDefaultBin);
        }
    }

    public WdPaperSize PaperSize
    {
        get => _pageSetup?.PaperSize.EnumConvert(WdPaperSize.wdPaper10x14) ?? WdPaperSize.wdPaper10x14;
        set
        {
            if (_pageSetup != null) _pageSetup.PaperSize = value.EnumConvert(MsWord.WdPaperSize.wdPaper10x14);
        }
    }

    public bool TwoPagesOnOne
    {
        get => _pageSetup?.TwoPagesOnOne ?? false;
        set
        {
            if (_pageSetup != null) _pageSetup.TwoPagesOnOne = value;
        }
    }

    #endregion

    #region 其他设置

    public int SuppressEndnotes
    {
        get => _pageSetup?.SuppressEndnotes ?? 0;
        set
        {
            if (_pageSetup != null) _pageSetup.SuppressEndnotes = value;
        }
    }

    public bool ShowGrid
    {
        get => _pageSetup?.ShowGrid ?? false;
        set
        {
            if (_pageSetup != null) _pageSetup.ShowGrid = value;
        }
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
            _pageSetup = null;
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