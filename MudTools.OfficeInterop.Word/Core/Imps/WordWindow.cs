//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// <see cref="IWordWindow"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.Window 对象。
/// </summary>
internal class WordWindow : IWordWindow
{
    private MsWord.Window _window;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordWindow"/> 类的新实例。
    /// </summary>
    /// <param name="window">原始的 Microsoft.Office.Interop.Word.Window 对象。</param>
    internal WordWindow(MsWord.Window window)
    {
        _window = window ?? throw new ArgumentNullException(nameof(window));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _window?.Application != null ? new WordApplication(_window.Application) : null;

    /// <inheritdoc/>
    public bool? DisplayVerticalScrollBar
    {
        get => _window?.DisplayVerticalScrollBar;
        set { if (_window != null) _window.DisplayVerticalScrollBar = value == true; }
    }

    /// <inheritdoc/>
    public bool? DisplayHorizontalScrollBar
    {
        get => _window?.DisplayHorizontalScrollBar;
        set { if (_window != null) _window.DisplayHorizontalScrollBar = value == true; }
    }

    /// <inheritdoc/>
    public IWordPanes? Panes => _window?.Panes != null ? new WordPanes(_window.Panes) : null;

    /// <inheritdoc/>
    public bool? Active
    {
        get => _window?.Active;
    }

    /// <inheritdoc/>
    public int? Hwnd => _window?.Hwnd;

    /// <inheritdoc/>
    public IWordDocument? Document => _window?.Document != null ? new WordDocument(_window.Document) : null;

    /// <inheritdoc/>
    public IWordView? View => _window?.View != null ? new WordView(_window.View) : null;


    /// <inheritdoc/>
    public IWordWindow? Next => _window?.Next != null ? new WordWindow(_window.Next) : null;

    /// <inheritdoc/>
    public IWordWindow? Previous => _window?.Previous != null ? new WordWindow(_window.Previous) : null;

    /// <inheritdoc/>
    public int? VerticalPercentScrolled
    {
        get => _window?.VerticalPercentScrolled;
        set { if (_window != null) _window.VerticalPercentScrolled = value ?? 0; }
    }

    /// <inheritdoc/>
    public int? HorizontalPercentScrolled
    {
        get => _window?.HorizontalPercentScrolled;
        set { if (_window != null) _window.HorizontalPercentScrolled = value ?? 0; }
    }

    /// <inheritdoc/>
    public int? Height
    {
        get => _window?.Height;
        set { if (_window != null) _window.Height = value ?? 0; }
    }

    /// <inheritdoc/>
    public int? Width
    {
        get => _window?.Width;
        set { if (_window != null) _window.Width = value ?? 0; }
    }

    /// <inheritdoc/>
    public string Caption
    {
        get => _window?.Caption;
        set { if (_window != null) _window.Caption = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public object Parent => _window?.Parent;

    /// <inheritdoc/>
    public int? Index => _window?.Index;

    /// <inheritdoc/>
    public int? Left
    {
        get => _window?.Left;
        set { if (_window != null) _window.Left = value ?? 0; }
    }

    /// <inheritdoc/>
    public int? Top
    {
        get => _window?.Top;
        set { if (_window != null) _window.Top = value ?? 0; }
    }

    /// <inheritdoc/>
    public bool? Visible
    {
        get => _window?.Visible;
        set { if (_window != null) _window.Visible = value ?? false; }
    }


    /// <inheritdoc/>
    public WdWindowType Type => _window != null ? (WdWindowType)_window.Type : WdWindowType.wdWindowDocument; // 默认值

    /// <inheritdoc/>
    public bool? WindowState
    {
        get => _window != null ? _window.WindowState == MsWord.WdWindowState.wdWindowStateMaximize : (bool?)null;
        set
        {
            if (_window != null)
            {
                _window.WindowState = value == true ?
                    MsWord.WdWindowState.wdWindowStateMaximize :
                    MsWord.WdWindowState.wdWindowStateNormal; // 或 wdWindowStateMinimize
            }
        }
    }

    /// <inheritdoc/>
    public int? SplitVertical
    {
        get => _window?.SplitVertical;
        set { if (_window != null) _window.SplitVertical = value ?? 0; }
    }

    #endregion // 属性实现

    #region 方法实现

    /// <inheritdoc/>
    public void Activate()
    {
        _window?.Activate();
    }


    /// <inheritdoc/>
    public void NewWindow()
    {
        _window?.NewWindow();
    }

    /// <inheritdoc/>
    public void PageScroll()
    {
        _window?.PageScroll();
    }

    /// <inheritdoc/>
    public void SetFocus()
    {
        _window?.SetFocus();
    }

    /// <inheritdoc/>
    public IWordRange? RangeFromPoint(int x, int y)
    {
        if (_window == null) return null;
        var obj = _window?.RangeFromPoint(x, y);
        if (obj != null && obj is MsWord.Range range)
            return new WordRange(range);
        return null;
    }

    public void GetPoint(out int ScreenPixelsLeft, out int ScreenPixelsTop, out int ScreenPixelsWidth, out int ScreenPixelsHeight, object obj)
    {
        _window.GetPoint(out ScreenPixelsLeft, out ScreenPixelsTop, out ScreenPixelsWidth, out ScreenPixelsHeight, obj);
    }


    /// <inheritdoc/>
    public void Close(WdSaveOptions saveChanges = WdSaveOptions.wdPromptToSaveChanges,
         bool routeDocument = false)
    {
        try
        {
            _window?.Close(
               (MsWord.WdSaveOptions)(int)saveChanges,
                routeDocument
            );
        }
        catch (COMException ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to close window: {ex.Message}");
            throw;
        }
    }

    /// <inheritdoc/>
    public void LargeScroll
        (int? down = null,
        int? up = null,
        int? toRight = null,
        int? toLeft = null)
    {
        object downObj = down ?? (object)missing;
        object upObj = up ?? (object)missing;
        object toRightObj = toRight ?? (object)missing;
        object toLeftObj = toLeft ?? (object)missing;

        _window?.LargeScroll(
            ref downObj,
            ref upObj,
            ref toRightObj,
            ref toLeftObj
        );
    }

    /// <inheritdoc/>
    public void ScrollIntoView(IWordRange range, bool? scrollToTopOfRange = null)
    {
        if (range == null) return;
        var comRange = (range as WordRange)?.InternalComObject;
        if (comRange == null) return;

        object scrollToTopObj = scrollToTopOfRange ?? (object)missing;
        _window?.ScrollIntoView(comRange, ref scrollToTopObj);
    }

    /// <inheritdoc/>
    public void PageScroll(int? pages = null, int? lines = null)
    {
        object pagesObj = pages ?? (object)missing;
        object linesObj = lines ?? (object)missing;
        _window?.PageScroll(ref pagesObj, ref linesObj);
    }


    /// <inheritdoc/>
    public void SmallScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null)
    {
        object downObj = down ?? (object)missing;
        object upObj = up ?? (object)missing;
        object toRightObj = toRight ?? (object)missing;
        object toLeftObj = toLeft ?? (object)missing;

        _window?.SmallScroll(
            ref downObj,
            ref upObj,
            ref toRightObj,
            ref toLeftObj
        );
    }

    /// <inheritdoc/>
    public void PrintOut(
        bool? background = null,
        bool? append = null,
        WdPrintOutRange? range = null,
        string outputFileName = null,
        object from = null,
        object to = null,
        WdPrintOutItem? item = null,
        int? copies = null,
        WdPrintOutPages? pages = null,
        bool? activePrinterMacGX = null,
        bool? manualDuplexPrint = null,
        bool? printDiskPromptForEachSheet = null,
        bool? collate = null,
        string fileName = null,
        bool? lineNumbers = null,
        int? numCopies = null)
    {
        // Prepare optional parameters for COM interop
        object backgroundObj = background ?? (object)missing;
        object appendObj = append ?? (object)missing;
        object rangeObj = range ?? (object)missing;
        object outputFileNameObj = outputFileName ?? (object)missing;
        object fromObj = from ?? (object)missing;
        object toObj = to ?? (object)missing;
        object itemObj = item ?? (object)missing;
        object copiesObj = copies ?? (object)missing;
        object pagesObj = pages ?? (object)missing;
        object activePrinterMacGXObj = activePrinterMacGX ?? (object)missing;
        object manualDuplexPrintObj = manualDuplexPrint ?? (object)missing;
        object printDiskPromptForEachSheetObj = printDiskPromptForEachSheet ?? (object)missing;
        object collateObj = collate ?? (object)missing;
        object fileNameObj = fileName ?? (object)missing;
        object lineNumbersObj = lineNumbers ?? (object)missing;
        object numCopiesObj = numCopies ?? (object)missing;

        try
        {
            _window?.Document?.PrintOut(
                ref backgroundObj,
                ref appendObj,
                ref rangeObj,
                ref outputFileNameObj,
                ref fromObj,
                ref toObj,
                ref itemObj,
                ref copiesObj,
                ref pagesObj,
                ref activePrinterMacGXObj,
                ref manualDuplexPrintObj,
                ref printDiskPromptForEachSheetObj,
                ref collateObj,
                ref fileNameObj,
                ref lineNumbersObj,
                ref numCopiesObj
            );
        }
        catch (COMException ex)
        {
            // Handle potential COM exceptions during print
            System.Diagnostics.Debug.WriteLine($"Failed to print window: {ex.Message}");
            throw; // Re-throw or handle as appropriate
        }
    }

    #endregion // 方法实现

    #region IDisposable 实现

    // 用于表示缺失的可选参数
    private static readonly object missing = System.Type.Missing;

    /// <summary>
    /// 释放由 <see cref="WordWindow"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_window != null)
                {
                    Marshal.ReleaseComObject(_window);
                    _window = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordWindow"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}