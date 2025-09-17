//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Window 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Window 对象的安全访问和资源管理
/// </summary>
internal class ExcelWindow : IExcelWindow
{
    private readonly MsExcel.Window _window;
    private bool _disposedValue;
    private double _savedHeight;
    private double _savedWidth;
    private double _savedLeft;
    private double _savedTop;
    private IExcelComSheet? _activeSheet;

    public IExcelRange? ActiveCell =>
        _window != null ? new ExcelRange(_window.ActiveCell) : null;

    public IExcelChart? ActiveChart =>
        _window != null ? new ExcelChart(_window.ActiveChart) : null;

    public IExcelPane? ActivePane =>
       _window != null ? new ExcelPane(_window.ActivePane) : null;

    public IExcelComSheet? ActiveSheet
    {
        get
        {
            if (_window == null) return null;
            if (_activeSheet != null) return _activeSheet;

            _activeSheet = Utils.CreateSheetObj(_window.ActiveSheet);
            return _activeSheet;
        }
    }

    public IExcelWorksheet? ActiveWorkSheet
    {
        get
        {
            if (ActiveSheet == null) return null;

            if (ActiveSheet is IExcelWorksheet worksheet)
                return worksheet;
            return null;
        }
    }


    public string Caption
    {
        get => _window.Caption?.ToString();
        set => _window.Caption = value;
    }

    public double Height
    {
        get => _window.Height;
        set => _window.Height = value;
    }

    public double Width
    {
        get => _window.Width;
        set => _window.Width = value;
    }

    public double Left
    {
        get => _window.Left;
        set => _window.Left = value;
    }

    public double Top
    {
        get => _window.Top;
        set => _window.Top = value;
    }

    /// <summary>
    /// 获取坐标轴所在的 Application 对象
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_window.Application);

    public XlWindowState WindowState
    {
        get => (XlWindowState)(int)_window.WindowState;
        set => _window.WindowState = (MsExcel.XlWindowState)(int)value;
    }

    public XlWindowView View
    {
        get => (XlWindowView)(int)_window.View;
        set => _window.View = (MsExcel.XlWindowView)(int)value;
    }

    public double Zoom
    {
        get => Convert.ToDouble(_window.Zoom);
        set => _window.Zoom = value;
    }

    public bool FreezePanes
    {
        get => _window.FreezePanes;
        set => _window.FreezePanes = value;
    }

    public int SplitRow
    {
        get => _window.SplitRow;
        set => _window.SplitRow = value;
    }

    public int SplitColumn
    {
        get => _window.SplitColumn;
        set => _window.SplitColumn = value;
    }

    public bool Split
    {
        get => _window.Split;
        set => _window.Split = value;
    }

    public IExcelRange? VisibleRange => _window != null ? new ExcelRange(_window.VisibleRange) : null;

    public IExcelSheetViews SheetViews => _window != null ? new ExcelSheetViews(_window.SheetViews) : null;


    /// <summary>
    /// 获取或设置是否显示网格线
    /// </summary>
    public bool DisplayGridlines
    {
        get => _window.DisplayGridlines;
        set => _window.DisplayGridlines = value;
    }

    /// <summary>
    /// 获取或设置是否显示行列标题
    /// </summary>
    public bool DisplayHeadings
    {
        get => _window.DisplayHeadings;
        set => _window.DisplayHeadings = value;
    }

    /// <summary>
    /// 获取或设置是否显示零值
    /// </summary>
    public bool DisplayZeros
    {
        get => _window.DisplayZeros;
        set => _window.DisplayZeros = value;
    }

    /// <summary>
    /// 获取或设置是否从右到左显示
    /// </summary>
    public bool DisplayRightToLeft
    {
        get => _window.DisplayRightToLeft;
        set => _window.DisplayRightToLeft = value;
    }

    /// <summary>
    /// 获取或设置是否显示公式
    /// </summary>
    public bool DisplayFormulas
    {
        get => _window.DisplayFormulas;
        set => _window.DisplayFormulas = value;
    }

    /// <summary>
    /// 获取或设置是否显示水平滚动条
    /// </summary>
    public bool DisplayHorizontalScrollBar
    {
        get => _window.DisplayHorizontalScrollBar;
        set => _window.DisplayHorizontalScrollBar = value;
    }

    /// <summary>
    /// 获取或设置是否显示垂直滚动条
    /// </summary>
    public bool DisplayVerticalScrollBar
    {
        get => _window.DisplayVerticalScrollBar;
        set => _window.DisplayVerticalScrollBar = value;
    }

    /// <summary>
    /// 获取或设置是否显示工作表标签
    /// </summary>
    public bool DisplayWorkbookTabs
    {
        get => _window.DisplayWorkbookTabs;
        set => _window.DisplayWorkbookTabs = value;
    }

    public bool DisplayRuler
    {
        get => _window.DisplayRuler;
        set => _window.DisplayRuler = value;
    }

    public bool AutoFilterDateGrouping
    {
        get => _window.AutoFilterDateGrouping;
        set => _window.AutoFilterDateGrouping = value;
    }

    public bool DisplayWhitespace
    {
        get => _window.DisplayWhitespace;
        set => _window.DisplayWhitespace = value;
    }

    /// <summary>
    /// 获取或设置当前垂直滚动位置（行号）
    /// </summary>
    public int ScrollRow
    {
        get => _window.ScrollRow;
        set => _window.ScrollRow = value;
    }

    /// <summary>
    /// 获取或设置当前水平滚动位置（列号）
    /// </summary>
    public int ScrollColumn
    {
        get => _window.ScrollColumn;
        set => _window.ScrollColumn = value;
    }

    /// <summary>
    /// 获取或设置图表对象是否可见
    /// </summary>
    public bool Visible
    {
        get => _window != null && _window.Visible;
        set
        {
            if (_window != null)
                _window.Visible = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示分级显示符号
    /// </summary>
    public bool DisplayOutline
    {
        get => _window.DisplayOutline;
        set => _window.DisplayOutline = value;
    }

    /// <summary>
    /// 获取或设置网格线颜色（RGB值）
    /// </summary>
    public int GridlineColor
    {
        get => _window.GridlineColor;
        set => _window.GridlineColor = value;
    }

    /// <summary>
    /// 获取或设置网格线颜色索引
    /// </summary>
    public XlColorIndex GridlineColorIndex
    {
        get => (XlColorIndex)_window.GridlineColorIndex;
        set => _window.GridlineColorIndex = (MsExcel.XlColorIndex)value;
    }

    /// <summary>
    /// 获取窗口句柄（HWND）
    /// </summary>
    public int Hwnd => _window.Hwnd;

    /// <summary>
    /// 获取窗口中的窗格集合
    /// </summary>
    public IExcelPanes Panes
    {
        get
        {
            return new ExcelPanes(_window.Panes);
        }
    }

    public IExcelRange? Selection
    {
        get
        {
            MsExcel.Range? range = _window.Selection as MsExcel.Range;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 获取当前选中的单元格区域
    /// </summary>
    public IExcelRange? RangeSelection
    {
        get
        {
            MsExcel.Range range = _window.RangeSelection;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 获取或设置水平拆分位置（像素）
    /// </summary>
    public double SplitHorizontal
    {
        get => _window.SplitHorizontal;
        set => _window.SplitHorizontal = value;
    }

    /// <summary>
    /// 获取或设置垂直拆分位置（像素）
    /// </summary>
    public double SplitVertical
    {
        get => _window.SplitVertical;
        set => _window.SplitVertical = value;
    }

    /// <summary>
    /// 获取或设置工作表标签区域占比
    /// </summary>
    public double TabRatio
    {
        get => _window.TabRatio;
        set => _window.TabRatio = value;
    }

    /// <summary>
    /// 获取窗口类型（工作表/图表）
    /// </summary>
    public XlWindowType Type => (XlWindowType)(int)_window.Type;

    /// <summary>
    /// 获取窗口可用高度（排除工具栏等）
    /// </summary>
    public double UsableHeight => _window.UsableHeight;

    /// <summary>
    /// 获取窗口可用宽度（排除工具栏等）
    /// </summary>
    public double UsableWidth => _window.UsableWidth;

    private IExcelSheets? _selectedSheets;

    /// <summary>
    /// 获取选中的工作表集合
    /// </summary>
    public IExcelSheets? SelectedSheets
    {
        get
        {
            if (_selectedSheets != null)
                return _selectedSheets;
            _selectedSheets = _window != null ? new ExcelSheets(_window.SelectedSheets) : null;
            return _selectedSheets;
        }
    }

    public object? Parent
    {
        get
        {
            if (_window == null || _window.Parent == null)
                return null;
            if (_window.Parent is MsExcel.Application app)
                return new ExcelApplication(app);

            if (_window.Parent is MsExcel.Workbook workbook)
                return new ExcelWorkbook(workbook);

            if (_window.Parent is MsExcel.Windows wins)
                return new ExcelWindows(wins, null);

            return _window.Parent;
        }
    }

    public IExcelWorkbook? ParentWorkbook
    {
        get
        {
            if (_window == null || _window.Parent == null)
                return null;
            if (_window.Parent is MsExcel.Workbook workbook)
                return new ExcelWorkbook(workbook);
            return null;
        }
    }

    public int Index => _window.Index;

    /// <summary>
    /// 获取窗口是否处于活动状态
    /// </summary>
    public bool IsActive
    {
        get
        {
            try
            {
                return _window.Application.ActiveWindow?.Hwnd == _window.Hwnd;
            }
            catch
            {
                return false;
            }
        }
    }

    internal ExcelWindow(MsExcel.Window window)
    {
        _window = window ?? throw new ArgumentNullException(nameof(window));
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _window.Activate();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate window.", ex);
        }
    }

    public void Close()
    {
        try
        {
            _window.Close();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close window.", ex);
        }
    }

    public void SelectRange(string rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("Range address cannot be null or empty", nameof(rangeAddress));

        try
        {
            if (ActiveSheet != null && ActiveSheet is IExcelWorksheet worksheet)
            {
                var range = worksheet?.Range(rangeAddress, null);
                range?.Select();
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to select range '{rangeAddress}'", ex);
        }
    }

    public object? RangeFromPoint(int x, int y)
    {
        var obj = _window?.RangeFromPoint(x, y);
        if (obj is MsExcel.Range range)
            return new ExcelRange(range);
        var control = Utils.CreateControl(obj, XlFormControl.xlDropDown);
        if (control != null)
            return control;
        control = Utils.CreateControl(obj, XlFormControl.xlListBox);
        if (control != null)
            return control;
        control = Utils.CreateControl(obj, XlFormControl.xlCheckBox);
        if (control != null)
            return control;
        control = Utils.CreateControl(obj, XlFormControl.xlEditBox);
        if (control != null)
            return control;
        return obj;
    }

    public void ScrollToRange(string rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("Range address cannot be null or empty", nameof(rangeAddress));

        try
        {
            if (ActiveSheet != null && ActiveSheet is IExcelWorksheet worksheet)
            {
                var range = worksheet?.Range(rangeAddress, null);
                if (range != null)
                {
                    // 确保区域可见
                    range.Activate();

                    // 滚动到区域位置
                    _window.ScrollRow = range.Row;
                    _window.ScrollColumn = range.Column;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to scroll to range '{rangeAddress}'", ex);
        }
    }

    public void Refresh()
    {
        try
        {
            _window.SmallScroll(); // 触发刷新
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh window.", ex);
        }
    }

    public void SaveLayout()
    {
        _savedHeight = Height;
        _savedWidth = Width;
        _savedLeft = Left;
        _savedTop = Top;
    }

    public void RestoreLayout()
    {
        Height = _savedHeight;
        Width = _savedWidth;
        Left = _savedLeft;
        Top = _savedTop;
    }

    #region 新增方法
    /// <summary>
    /// 创建当前窗口的新实例
    /// </summary>
    /// <returns>新窗口对象</returns>
    public IExcelWindow NewWindow()
    {
        var newWindow = _window.NewWindow();
        return new ExcelWindow(newWindow);
    }

    /// <summary>
    /// 大范围滚动窗口内容
    /// </summary>
    /// <param name="down">向下滚动页数</param>
    /// <param name="up">向上滚动页数</param>
    /// <param name="right">向右滚动页数</param>
    /// <param name="left">向左滚动页数</param>
    public void LargeScroll(int? down = 0, int? up = 0, int? right = 0, int? left = 0)
    {
        _window.LargeScroll(
            Down: down.ComArgsVal(),
            Up: up.ComArgsVal(),
            ToRight: right.ComArgsVal(),
            ToLeft: left.ComArgsVal()
        );
    }

    /// <summary>
    /// 小范围滚动窗口内容
    /// </summary>
    /// <param name="down">向下滚动行数</param>
    /// <param name="up">向上滚动行数</param>
    /// <param name="right">向右滚动列数</param>
    /// <param name="left">向左滚动列数</param>
    public void SmallScroll(int? down = 0, int? up = 0, int? right = 0, int? left = 0)
    {
        _window.SmallScroll(
            Down: down.ComArgsVal(),
            Up: up.ComArgsVal(),
            ToRight: right.ComArgsVal(),
            ToLeft: left.ComArgsVal()
        );
    }

    /// <summary>
    /// 将水平坐标点转换为屏幕像素值
    /// </summary>
    /// <param name="points">点坐标</param>
    /// <returns>像素值</returns>
    public int PointsToScreenPixelsX(int points)
    {
        return _window.PointsToScreenPixelsX(points);
    }

    /// <summary>
    /// 将垂直坐标点转换为屏幕像素值
    /// </summary>
    /// <param name="points">点坐标</param>
    /// <returns>像素值</returns>
    public int PointsToScreenPixelsY(int points)
    {
        return _window.PointsToScreenPixelsY(points);
    }

    /// <summary>
    /// 封装Excel窗口打印方法
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="preview">是否预览</param>
    /// <param name="activePrinter">打印机名称</param>
    /// <param name="printToFile">是否打印到文件</param>
    /// <param name="collate">是否逐份打印</param>
    public void PrintOut(
        int copies = 1,
        bool preview = false,
        string? activePrinter = null,
        bool printToFile = false,
        bool collate = true)
    {
        object missing = System.Type.Missing;

        _window.PrintOut(
            From: missing,
            To: missing,
            Copies: copies,
            Preview: preview,
            ActivePrinter: activePrinter ?? missing,
            PrintToFile: printToFile,
            Collate: collate,
            PrToFileName: missing
        );
    }


    /// <summary>
    /// 打印窗口内容
    /// </summary>
    /// <param name="preview">是否预览</param>
    public void PrintOut(bool preview = false)
    {
        if (preview)
        {
            _window.PrintPreview();
        }
        else
        {
            _window.PrintOut(
                From: System.Type.Missing,
                To: System.Type.Missing,
                Copies: System.Type.Missing,
                Preview: System.Type.Missing,
                ActivePrinter: System.Type.Missing,
                PrintToFile: System.Type.Missing,
                Collate: System.Type.Missing,
                PrToFileName: System.Type.Missing
            );
        }
    }

    /// <summary>
    /// 执行打印预览
    /// </summary>
    public void PrintPreview()
    {
        try
        {
            // 执行原生打印预览
            _window.PrintPreview(EnableChanges: false);
        }
        catch (Exception ex)
        {
            throw new ApplicationException("打印预览失败", ex);
        }
    }
    #endregion
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _activeSheet?.Dispose();
            _selectedSheets?.Dispose();
            _selectedSheets = null;
        }
        _activeSheet = null;
        _selectedSheets = null;
        _selectedSheets = null;
        _disposedValue = true;
    }
    ~ExcelWindow()
    {
        Dispose(false);
    }
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
