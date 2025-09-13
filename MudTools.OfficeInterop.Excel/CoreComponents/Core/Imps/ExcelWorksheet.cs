//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using System.Drawing;
using System.Reflection;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Worksheet 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Worksheet 对象的安全访问和资源管理
/// </summary>
internal partial class ExcelWorksheet : IExcelWorksheet
{
    private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

    /// <summary>
    /// 底层的 COM Worksheet 对象
    /// </summary>
    private MsExcel.Worksheet _worksheet;

    internal MsExcel.Worksheet Worksheet => _worksheet;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放
    /// <summary>
    /// 初始化 ExcelWorksheet 实例
    /// </summary>
    /// <param name="worksheet">底层的 COM Worksheet 对象</param>
    internal ExcelWorksheet(MsExcel.Worksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _docEvents_Event = worksheet;
        InitializeEvents();
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放子COM组件
                _listObjects?.Dispose();
                _names?.Dispose();
                _vPageBreaks?.Dispose();
                _hPageBreaks?.Dispose();
                _cells?.Dispose();
                _circularReference?.Dispose();
                _sort?.Dispose();
                _pageSetup?.Dispose();
                _shapes?.Dispose();
                _pictures?.Dispose();
                _hyperlinks?.Dispose();
                _comments?.Dispose();
                _usedRange?.Dispose();
                _allRange?.Dispose();
                _columns?.Dispose();
                _rows?.Dispose();

                DisConnectEvent();

                // 释放底层COM对象
                if (_worksheet != null)
                    Marshal.ReleaseComObject(_worksheet);

                if (_docEvents_Event != null)
                    Marshal.ReleaseComObject(_docEvents_Event);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
        }

        // 清理事件处理程序引用
        _circularReference = null;
        _listObjects = null;
        _hPageBreaks = null;
        _vPageBreaks = null;
        _names = null;
        _cells = null;
        _sort = null;
        _columns = null;
        _rows = null;
        _change = null;
        _selectionChange = null;
        _sheetActivate = null;
        _sheetDeactivate = null;
        _beforeDoubleClick = null;
        _beforeRightClick = null;
        _sheetCalculate = null;
        _worksheet = null;
        _docEvents_Event = null;
        _disposedValue = true;
    }

    ~ExcelWorksheet()
    {
        Dispose(false);
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置工作表的名称
    /// </summary>
    public string Name
    {
        get => _worksheet?.Name?.ToString();
        set
        {
            if (_worksheet != null && value != null)
                _worksheet.Name = value;
        }
    }

    private IExcelNames? _names = null;

    public IExcelNames Names
    {
        get
        {
            if (_names != null)
                return _names;
            _names = new ExcelNames(_worksheet.Names);
            return _names;
        }
    }
    private IExcelVPageBreaks? _vPageBreaks = null;

    public IExcelVPageBreaks VPageBreaks
    {
        get
        {
            if (_vPageBreaks != null)
                return _vPageBreaks;
            _vPageBreaks = new ExcelVPageBreaks(_worksheet.VPageBreaks);
            return _vPageBreaks;
        }
    }

    private IExcelHPageBreaks? _hPageBreaks = null;

    public IExcelHPageBreaks HPageBreaks
    {
        get
        {
            if (_hPageBreaks != null)
                return _hPageBreaks;
            _hPageBreaks = new ExcelHPageBreaks(_worksheet.HPageBreaks);
            return _hPageBreaks;
        }
    }

    private IExcelListObjects _listObjects;

    public IExcelListObjects ListObjects
    {
        get
        {
            if (_listObjects != null)
                return _listObjects;
            _listObjects = new ExcelListObjects(_worksheet.ListObjects);
            return _listObjects;
        }
    }

    private IExcelCells? _circularReference = null;
    public IExcelCells CircularReference
    {
        get
        {
            if (_circularReference != null)
                return _circularReference;
            _circularReference = new ExcelCells(_worksheet.CircularReference);
            return _circularReference;
        }
    }

    private IExcelCells? _cells = null;
    public IExcelCells Cells
    {
        get
        {
            if (_cells != null)
                return _cells;
            _cells = new ExcelCells(_worksheet.Cells);
            return _cells;
        }
    }
    private IExcelSort? _sort = null;
    public IExcelSort Sort
    {
        get
        {
            if (_sort != null)
                return _sort;
            _sort = new ExcelSort(_worksheet.Sort);
            return _sort;
        }
    }

    private Color? _tabColor;

    /// <summary>
    /// 获取或设置工作表标签颜色
    /// </summary>
    public Color TabColor
    {
        get
        {
            if (_tabColor == null && (int)_worksheet.Tab.Color != 0)
            {
                _tabColor = ColorTranslator.FromOle((int)_worksheet.Tab.Color);
            }
            return _tabColor ?? Color.Empty;
        }
        set
        {
            _tabColor = value;
            _worksheet.Tab.Color = value.ToArgb();
        }
    }

    public bool ProtectDrawingObjects => _worksheet.ProtectDrawingObjects;

    public bool ProtectScenarios => _worksheet.ProtectScenarios;

    public bool ProtectionMode => _worksheet.ProtectionMode;

    public bool TransitionExpEval
    {
        get => _worksheet.TransitionExpEval;
        set => _worksheet.TransitionExpEval = value;
    }

    /// <summary>
    /// 获取或设置标准列宽
    /// </summary>
    public double StandardWidth
    {
        get => _worksheet.StandardWidth;
        set => _worksheet.StandardWidth = value;
    }

    /// <summary>
    /// 获取大纲（分级显示）设置对象
    /// </summary>
    public IExcelOutline Outline => new ExcelOutline(_worksheet.Outline);

    /// <summary>
    /// 获取或设置自动筛选模式状态
    /// </summary>
    public bool AutoFilterMode
    {
        get => _worksheet.AutoFilterMode;
        set => _worksheet.AutoFilterMode = value;
    }

    public bool DisplayPageBreaks
    {
        get => _worksheet.DisplayPageBreaks;
        set => _worksheet.DisplayPageBreaks = value;
    }

    public bool ProtectContents
    {
        get => _worksheet.ProtectContents;
    }

    /// <summary>
    /// 获取工作表当前是否处于筛选模式
    /// </summary>
    public bool FilterMode => _worksheet.FilterMode;

    public bool EnableOutlining
    {
        get => _worksheet.EnableOutlining;
        set => _worksheet.EnableOutlining = value;
    }

    public bool EnablePivotTable
    {
        get => _worksheet.EnablePivotTable;
        set => _worksheet.EnablePivotTable = value;
    }

    public string OnCalculate
    {
        get => _worksheet.OnCalculate;
        set => _worksheet.OnCalculate = value;
    }

    public string OnData
    {
        get => _worksheet.OnData;
        set => _worksheet.OnData = value;
    }

    public string OnDoubleClick
    {
        get => _worksheet.OnDoubleClick;
        set => _worksheet.OnDoubleClick = value;
    }

    public bool DisplayAutomaticPageBreaks
    {
        get => _worksheet.DisplayAutomaticPageBreaks;
        set => _worksheet.DisplayAutomaticPageBreaks = value;
    }

    /// <summary>
    /// 获取工作表类型
    /// </summary>
    public XlSheetType Type => (XlSheetType)_worksheet.Type;

    public XlEnableSelection EnableSelection
    {
        get => (XlEnableSelection)_worksheet.EnableSelection;
        set
        {
            if (_worksheet != null)
                _worksheet.EnableSelection = (MsExcel.XlEnableSelection)(int)value;
        }
    }

    /// <summary>
    /// 获取工作表的索引位置
    /// </summary>
    public int Index => _worksheet?.Index ?? 0;

    /// <summary>
    /// 获取或设置工作表是否可见
    /// </summary>
    public XlSheetVisibility Visible
    {
        get => _worksheet != null ? (XlSheetVisibility)_worksheet.Visible : XlSheetVisibility.xlSheetHidden;
        set
        {
            if (_worksheet != null)
                _worksheet.Visible = (MsExcel.XlSheetVisibility)value;
        }
    }

    public bool IsVisible
    {
        get => _worksheet != null && _worksheet.Visible == MsExcel.XlSheetVisibility.xlSheetVisible;
        set
        {
            if (_worksheet != null)
                _worksheet.Visible = value ? (MsExcel.XlSheetVisibility.xlSheetVisible) : (MsExcel.XlSheetVisibility.xlSheetHidden);
        }
    }

    /// <summary>
    /// 获取工作表是否被保护
    /// </summary>
    public bool IsProtected => _worksheet != null && _worksheet.ProtectContents;

    /// <summary>
    /// 获取工作表的代码名称
    /// </summary>
    public string? CodeName => _worksheet?.CodeName;

    /// <summary>
    /// 获取工作表所在的父对象。
    /// 对于标准的 Excel 工作表，其父对象是它所属的工作簿 (Workbook)。
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_worksheet?.Parent == null)
            {
                return null;
            }
            if (_worksheet.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            log?.Warn($"Unexpected Parent type for worksheet: {_worksheet.Parent.GetType().FullName}");
            return _worksheet.Parent;
        }
    }

    public string? ParentName
    {
        get
        {
            if (_worksheet?.Parent == null)
            {
                return null;
            }
            if (_worksheet.Parent is MsExcel.Workbook workbook)
            {
                return workbook.Name;
            }
            if (_worksheet.Parent is MsExcel.Worksheet worksheet)
            {
                return worksheet.Name;
            }
            return null;
        }
    }

    public IExcelWorkbook? ParentWorkbook
    {
        get
        {
            if (_worksheet?.Parent == null)
            {
                return null;
            }
            if (_worksheet.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            log?.Warn($"Unexpected Parent type for worksheet: {_worksheet.Parent.GetType().FullName}");
            return null;
        }
    }

    /// <summary>
    /// 获取工作表所在的Application对象
    /// </summary>
    public IExcelApplication Application
    {
        get
        {
            var application = _worksheet?.Application as MsExcel.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    /// <summary>
    /// 获取工作表的下一个工作表
    /// </summary>
    public IExcelWorksheet Next
    {
        get
        {
            try
            {
                var nextSheet = _worksheet.Next as MsExcel.Worksheet;
                return nextSheet != null ? new ExcelWorksheet(nextSheet) : null;
            }
            catch
            {
                return null;
            }
        }
    }


    /// <summary>
    /// 获取工作表的上一个工作表
    /// </summary>
    public IExcelWorksheet Previous
    {
        get
        {
            try
            {
                var previousSheet = _worksheet.Previous as MsExcel.Worksheet;
                return previousSheet != null ? new ExcelWorksheet(previousSheet) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 区域访问

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    public IExcelRange? Range(object? cell1, object? cell2 = null)
    {
        if (_worksheet == null) return null;

        try
        {
            if (cell1 is ExcelRange range1)
                cell1 = range1.InternalRange;
            if (cell2 is ExcelRange range2)
                cell2 = range2.InternalRange;

            cell1 ??= System.Type.Missing;
            cell2 ??= System.Type.Missing;

            var range = _worksheet.Range[cell1, cell2]; ;
            return range != null ? new ExcelRange(range) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    public IExcelRange? this[int row, int column]
    {
        get
        {
            if (_worksheet == null) return null;

            try
            {
                MsExcel.Range? range = _worksheet.Cells[row, column] as MsExcel.Range;
                return range != null ? new ExcelRange(range) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public IExcelRange? this[string address]
    {
        get
        {
            if (_worksheet == null) return null;
            try
            {
                var range = _worksheet.Range[address];
                return range != null ? new ExcelRange(range) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public IExcelRange? this[string begin, string end] => this[$"{begin}:{end}"];


    /// <summary>
    /// 获取工作表中指定行的区域对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>行区域对象</returns>
    public IExcelRange GetRow(int row)
    {
        if (_worksheet == null) return null;

        try
        {
            var range = _worksheet.Rows[row] as MsExcel.Range;
            return range != null ? new ExcelRange(range) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取工作表中指定列的区域对象
    /// </summary>
    /// <param name="column">列号</param>
    /// <returns>列区域对象</returns>
    public IExcelRange GetColumn(int column)
    {
        if (_worksheet == null) return null;

        try
        {
            var range = _worksheet.Columns[column] as MsExcel.Range;
            return range != null ? new ExcelRange(range) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 已使用区域缓存
    /// </summary>
    private IExcelRange _usedRange;

    /// <summary>
    /// 获取工作表的已使用区域
    /// </summary>
    public IExcelRange UsedRange => _usedRange ?? (_usedRange = new ExcelRange(_worksheet?.UsedRange));

    /// <summary>
    /// 整个工作表区域缓存
    /// </summary>
    private IExcelRange _allRange;

    /// <summary>
    /// 获取工作表的整个区域
    /// </summary>
    public IExcelRange AllRange => _allRange ?? (_allRange = new ExcelRange(_worksheet?.Cells));


    private IExcelRange? _rows;
    private IExcelRange? _columns;

    public IExcelRange? Rows
    {
        get
        {
            if (_rows != null)
                return _rows;
            if (_worksheet == null)
                return null;
            _rows ??= new ExcelRange(_worksheet.Rows);
            return _rows;
        }
    }

    public IExcelRange? Columns
    {
        get
        {
            if (_columns != null)
                return _columns;
            if (_worksheet == null)
                return null;
            _columns ??= new ExcelRange(_worksheet.Columns);
            return _columns;
        }
    }
    #endregion

    #region 页面设置

    /// <summary>
    /// 页面设置对象缓存
    /// </summary>
    private IExcelPageSetup _pageSetup;

    /// <summary>
    /// 获取工作表的页面设置对象
    /// </summary>
    public IExcelPageSetup PageSetup => _pageSetup ??= new ExcelPageSetup(_worksheet?.PageSetup);

    #endregion

    #region 形状和图表

    /// <summary>
    /// 形状集合缓存
    /// </summary>
    private IExcelShapes _shapes;

    /// <summary>
    /// 获取工作表的形状集合
    /// </summary>
    public IExcelShapes Shapes => _shapes ??= new ExcelShapes(_worksheet?.Shapes);

    /// <summary>
    /// 图片集合缓存
    /// </summary>
    private IExcelPictures _pictures;

    /// <summary>
    /// 获取工作表的图片集合
    /// </summary>
    public IExcelPictures Pictures => _pictures ??= new ExcelPictures(_worksheet?.Pictures() as MsExcel.Pictures);

    /// <summary>
    /// 评论集合缓存
    /// </summary>
    private IExcelComments _comments;

    /// <summary>
    /// 获取工作表的评论集合
    /// </summary>
    public IExcelComments Comments => _comments ?? (_comments = new ExcelComments(_worksheet?.Comments));

    /// <summary>
    /// 超链接集合缓存
    /// </summary>
    private IExcelHyperlinks _hyperlinks;

    /// <summary>
    /// 获取工作表的超链接集合
    /// </summary>
    public IExcelHyperlinks Hyperlinks => _hyperlinks ?? (_hyperlinks = new ExcelHyperlinks(_worksheet?.Hyperlinks));

    #endregion

    #region 数据操作

    /// <summary>
    /// 获取或设置工作表的默认行高
    /// </summary>
    public double DefaultRowHeight
    {
        get => (double)(_worksheet?.Rows?.RowHeight ?? 0);
        set
        {
            if (_worksheet?.Rows != null)
                _worksheet.Rows.RowHeight = value;
        }
    }

    /// <summary>
    /// 获取或设置工作表的默认列宽
    /// </summary>
    public double DefaultColumnWidth
    {
        get => (double)(_worksheet?.Columns?.ColumnWidth ?? 0);
        set
        {
            if (_worksheet?.Columns != null)
                _worksheet.Columns.ColumnWidth = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用自动筛选
    /// </summary>
    public bool EnableAutoFilter
    {
        get => _worksheet != null && _worksheet.EnableAutoFilter;
        set
        {
            if (_worksheet != null)
                _worksheet.EnableAutoFilter = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用计算器
    /// </summary>
    public bool EnableCalculation
    {
        get => _worksheet != null && _worksheet.EnableCalculation;
        set
        {
            if (_worksheet != null)
                _worksheet.EnableCalculation = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示页面布局
    /// </summary>
    public bool DisplayPageLayout
    {
        get => _worksheet != null && _worksheet.DisplayPageBreaks;
        set
        {
            if (_worksheet != null)
                _worksheet.DisplayPageBreaks = value;
        }
    }
    #endregion

    #region 保护和安全

    /// <summary>
    /// 保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="drawingObjects">是否保护图形对象</param>
    /// <param name="contents">是否保护内容</param>
    /// <param name="scenarios">是否保护方案</param>
    /// <param name="userInterfaceOnly">是否仅保护用户界面</param>
    /// <param name="allowFormattingCells">是否允许格式化单元格</param>
    /// <param name="allowFormattingColumns">是否允许格式化列</param>
    /// <param name="allowFormattingRows">是否允许格式化行</param>
    /// <param name="allowInsertingColumns">是否允许插入列</param>
    /// <param name="allowInsertingRows">是否允许插入行</param>
    /// <param name="allowInsertingHyperlinks">是否允许插入超链接</param>
    /// <param name="allowDeletingColumns">是否允许删除列</param>
    /// <param name="allowDeletingRows">是否允许删除行</param>
    /// <param name="allowSorting">是否允许排序</param>
    /// <param name="allowFiltering">是否允许筛选</param>
    /// <param name="allowUsingPivotTables">是否允许使用透视表</param>
    public void Protect(string password = "", bool drawingObjects = true, bool contents = true,
                       bool scenarios = true, bool userInterfaceOnly = false,
                       bool allowFormattingCells = true, bool allowFormattingColumns = true,
                       bool allowFormattingRows = true, bool allowInsertingColumns = true,
                       bool allowInsertingRows = true, bool allowInsertingHyperlinks = true,
                       bool allowDeletingColumns = true, bool allowDeletingRows = true,
                       bool allowSorting = true, bool allowFiltering = true,
                       bool allowUsingPivotTables = true)
    {
        _worksheet?.Protect(password, drawingObjects, contents, scenarios, userInterfaceOnly,
                           allowFormattingCells, allowFormattingColumns, allowFormattingRows,
                           allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks,
                           allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering,
                           allowUsingPivotTables);
    }

    /// <summary>
    /// 保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="drawingObjects">是否保护图形对象</param>
    /// <param name="contents">是否保护内容</param>
    /// <param name="scenarios">是否保护方案</param>
    /// <param name="userInterfaceOnly">是否仅保护用户界面</param>
    public void Protect(string? password = null, bool? drawingObjects = null,
        bool? contents = null, bool? scenarios = null, bool? userInterfaceOnly = null)
    {
        _worksheet?.Protect(
          Password: password.ComArgsVal(),
          DrawingObjects: drawingObjects.ComArgsVal(),
          Contents: contents.ComArgsVal(),
          Scenarios: scenarios.ComArgsVal(),
          UserInterfaceOnly: userInterfaceOnly.ComArgsVal());
    }

    /// <summary>
    /// 取消保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public void Unprotect(string password = "")
    {
        _worksheet?.Unprotect(password);
    }

    #endregion

    #region 操作方法

    public object? OLEObjects(int? index = null)
    {
        if (index != null)
            return _worksheet.OLEObjects(index);
        return _worksheet?.OLEObjects();
    }

    /// <summary>
    /// 激活工作表
    /// </summary>
    public void Activate()
    {
        _worksheet?.Activate();
    }

    /// <summary>
    /// 选择工作表
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        _worksheet?.Select(replace);
    }

    /// <summary>
    /// 粘贴剪贴板内容到指定单元格
    /// </summary>
    /// <param name="destinationCell">目标单元格</param>
    public void Paste(IExcelRange destinationCell, bool? link = null)
    {
        if (_worksheet == null) return;
        object dest = System.Type.Missing;
        if (destinationCell != null)
            dest = ((ExcelRange)destinationCell).InternalRange;
        _worksheet.Paste(dest, link.ComArgsVal());
    }

    /// <summary>
    /// 粘贴剪贴板内容到指定区域
    /// </summary>
    /// <param name="startRow">起始行</param>
    /// <param name="startColumn">起始列</param>
    public void PasteToPosition(int startRow, int startColumn)
    {
        try
        {
            var destinationRange = _worksheet.Cells[startRow, startColumn];
            _worksheet.Paste(destinationRange);
        }
        catch (Exception ex)
        {
            throw new Exception($"粘贴失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 通用特殊粘贴方法
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    /// <param name="pasteType">粘贴类型</param>
    /// <param name="skipBlanks">是否跳过空白单元格</param>
    /// <param name="transpose">是否转置</param>
    public void PasteSpecial(IExcelRange destinationRange,
                           XlPasteType pasteType = XlPasteType.xlPasteAll,
                           bool skipBlanks = false,
                           bool transpose = false)
    {
        if (destinationRange == null)
            return;

        destinationRange.PasteSpecial(
                 paste: pasteType,
                 operation: XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                 skipBlanks: skipBlanks,
                 transpose: transpose
             );
    }

    /// <summary>
    /// 特殊粘贴 - 只粘贴值
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    public void PasteValues(IExcelRange destinationRange)
    {
        PasteSpecial(destinationRange, XlPasteType.xlPasteValues);
    }

    /// <summary>
    /// 特殊粘贴 - 只粘贴格式
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    public void PasteFormats(IExcelRange destinationRange)
    {
        PasteSpecial(destinationRange, XlPasteType.xlPasteFormats);
    }

    /// <summary>
    /// 特殊粘贴 - 粘贴公式
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    public void PasteFormulas(IExcelRange destinationRange)
    {
        PasteSpecial(destinationRange, XlPasteType.xlPasteFormulas);
    }

    /// <summary>
    /// 特殊粘贴 - 粘贴列宽
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    public void PasteColumnWidths(IExcelRange destinationRange)
    {
        PasteSpecial(destinationRange, XlPasteType.xlPasteColumnWidths);
    }

    /// <summary>
    /// 特殊粘贴 - 执行计算操作
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    /// <param name="operation">计算操作类型</param>
    public void PasteWithOperation(IExcelRange destinationRange, XlPasteSpecialOperation operation)
    {
        if (destinationRange == null)
            return;

        destinationRange.PasteSpecial(
                paste: XlPasteType.xlPasteAll,
                operation: operation,
                skipBlanks: false,
                transpose: false
            );
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="before">复制到指定工作表之前</param>
    /// <param name="after">复制到指定工作表之后</param>
    public void Copy(IExcelWorksheet? before = null, IExcelWorksheet? after = null)
    {
        if (_worksheet == null) return;

        _worksheet.Copy(
            before is ExcelWorksheet beforeSheet ? beforeSheet._worksheet : System.Type.Missing,
            after is ExcelWorksheet afterSheet ? afterSheet._worksheet : System.Type.Missing
        );
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    public void Copy()
    {
        _worksheet.Copy();
    }

    /// <summary>
    /// 移动工作表
    /// </summary>
    /// <param name="before">移动到指定工作表之前</param>
    /// <param name="after">移动到指定工作表之后</param>
    public void Move(IExcelWorksheet? before = null, IExcelWorksheet? after = null)
    {
        if (_worksheet == null) return;

        _worksheet.Move(
            before is ExcelWorksheet beforeSheet ? beforeSheet._worksheet : System.Type.Missing,
            after is ExcelWorksheet afterSheet ? afterSheet._worksheet : System.Type.Missing
        );
    }

    public IExcelPivotTables? PivotTables()
    {
        var pivotTables = _worksheet?.PivotTables() as MsExcel.PivotTables;
        if (pivotTables == null)
            return null;
        return new ExcelPivotTables(pivotTables);
    }

    public IExcelPivotTable? PivotTables(int index)
    {
        var pivotTable = _worksheet?.PivotTables(index) as MsExcel.PivotTable;
        if (pivotTable == null)
            return null;
        return new ExcelPivotTable(pivotTable);
    }

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    public IExcelChartObjects? ChartObjects()
    {
        var chartObjects = _worksheet?.ChartObjects() as MsExcel.ChartObjects;
        if (chartObjects == null)
            return null;
        return new ExcelChartObjects(chartObjects);
    }

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    public IExcelChartObject ChartObjects(int index)
    {
        var chartObject = _worksheet.ChartObjects(index) as MsExcel.ChartObject;
        return new ExcelChartObject(chartObject);
    }

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    public IExcelChartObject ChartObjects(string name)
    {
        var chartObject = _worksheet.ChartObjects(name) as MsExcel.ChartObject;
        return new ExcelChartObject(chartObject);
    }

    public void ExportAsFixedFormat(
        XlFixedFormatType Type,
        string Filename,
        object? Quality = null,
        object? IncludeDocProperties = null,
        object? IgnorePrintAreas = null,
        object? From = null,
        object? To = null,
        object? OpenAfterPublish = null,
        object? FixedFormatExtClassPtr = null)
    {
        _worksheet.ExportAsFixedFormat(
            (MsExcel.XlFixedFormatType)Type,
            Filename,
            Quality ?? System.Type.Missing,
            IncludeDocProperties ?? System.Type.Missing,
            IgnorePrintAreas ?? System.Type.Missing,
            From ?? System.Type.Missing,
            To ?? System.Type.Missing,
            OpenAfterPublish ?? System.Type.Missing,
            FixedFormatExtClassPtr ?? System.Type.Missing);
    }

    /// <summary>
    /// 删除工作表
    /// </summary>
    public void Delete()
    {
        _worksheet?.Delete();
    }


    /// <summary>
    /// 打印工作表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    public void PrintOut(bool preview = false)
    {
        if (_worksheet == null) return;

        if (preview)
        {
            _worksheet.PrintPreview();
        }
        else
        {
            _worksheet.PrintOut(
               System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                 System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing
            );
        }
    }

    /// <summary>
    /// 重命名工作表
    /// </summary>
    /// <param name="newName">新名称</param>
    public void Rename(string newName)
    {
        if (_worksheet != null && !string.IsNullOrEmpty(newName))
            _worksheet.Name = newName;
    }

    /// <summary>
    /// 在指定位置创建数据透视表
    /// </summary>
    /// <param name="sourceRange">数据源范围地址（如 "A1:D100"）</param>
    /// <param name="targetCell">目标位置左上角单元格地址（如 "F1"）</param>
    /// <param name="tableName">数据透视表名称</param>
    /// <returns>创建的数据透视表包装对象</returns>
    public ExcelPivotTable CreatePivotTable(string sourceRange, string targetCell, string tableName)
    {
        if (string.IsNullOrWhiteSpace(sourceRange))
            throw new ArgumentException("数据源范围不能为空");
        if (string.IsNullOrWhiteSpace(targetCell))
            throw new ArgumentException("目标位置不能为空");
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("数据透视表名称不能为空");

        try
        {
            // 获取数据源范围
            var source = _worksheet.Range[sourceRange];

            var workbook = _worksheet.Parent as MsExcel.Workbook;

            // 创建数据透视表缓存
            var pivotCache = workbook.PivotCaches().Create(
                SourceType: MsExcel.XlPivotTableSourceType.xlDatabase,
                SourceData: source);

            // 创建数据透视表
            var pivotTable = pivotCache.CreatePivotTable(
                TableDestination: _worksheet.Range[targetCell],
                TableName: tableName);

            // 释放中间COM对象
            Marshal.ReleaseComObject(source);
            Marshal.ReleaseComObject(pivotCache);
            Marshal.ReleaseComObject(workbook);

            return new ExcelPivotTable(pivotTable);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("创建数据透视表失败", ex);
        }
    }
    #endregion

    #region 高级功能
    /// <summary>
    /// 在活动工作表中创建指定范围
    /// </summary>
    public IExcelRange CreateRange(string address)
    {
        if (_worksheet != null)
        {
            return new ExcelRange(_worksheet.Range[address]);
        }
        return null;
    }

    /// <summary>
    /// 在活动工作表中获取指定行
    /// </summary>
    public IExcelRows GetRows(int startRow, int endRow = -1)
    {
        if (_worksheet == null)
            return null;

        if (endRow < startRow) endRow = startRow;
        var rang = _worksheet.Rows[$"{startRow}:{endRow}"] as MsExcel.Range;
        return new ExcelRange(rang);
    }

    /// <summary>
    /// 在活动工作表中获取指定列
    /// </summary>
    public IExcelColumns GetColumns(string startColumn, string endColumn = "")
    {
        if (_worksheet == null)
            return null;

        if (string.IsNullOrEmpty(endColumn)) endColumn = startColumn;
        var rang = _worksheet.Columns[$"{startColumn}:{endColumn}"] as MsExcel.Range;
        return new ExcelRange(rang);
    }



    /// <summary>
    /// 将工作表另存为xlsx文件。
    /// </summary>
    /// <param name="filePath"></param>
    public void SaveAs(string filePath)
    {
        _worksheet?.SaveAs(filePath);
    }

    public IExcelAutoFilter AutoFilter
    {
        get
        {
            return new ExcelAutoFilter(_worksheet.AutoFilter);
        }
    }


    public void ResetAllPageBreaks()
    {
        _worksheet?.ResetAllPageBreaks();
    }

    /// <summary>
    /// 计算工作表中的所有公式
    /// </summary>
    public void Calculate()
    {
        _worksheet?.Calculate();
    }

    /// <summary>
    /// 重新计算工作表
    /// </summary>
    public void Recalculate()
    {
        _worksheet?.Calculate();
    }

    /// <summary>
    /// 清除工作表内容
    /// </summary>
    public void Clear()
    {
        _worksheet?.UsedRange?.ClearContents();
    }

    /// <summary>
    /// 清除工作表格式
    /// </summary>
    public void ClearFormats()
    {
        _worksheet?.UsedRange?.ClearFormats();
    }

    /// <summary>
    /// 清除工作表内容和格式
    /// </summary>
    public void ClearAll()
    {
        _worksheet?.UsedRange?.Clear();
    }

    /// <summary>
    /// 清除工作表注释
    /// </summary>
    public void ClearComments()
    {
        _worksheet?.UsedRange?.ClearComments();
    }

    /// <summary>
    /// 清除工作表超链接
    /// </summary>
    public void ClearHyperlinks()
    {
        _worksheet?.Hyperlinks?.Delete();
    }

    /// <summary>
    /// 自动调整列宽
    /// </summary>
    public void AutoFitColumns()
    {
        _worksheet?.UsedRange?.EntireColumn?.AutoFit();
    }

    /// <summary>
    /// 自动调整行高
    /// </summary>
    public void AutoFitRows()
    {
        _worksheet?.UsedRange?.EntireRow?.AutoFit();
    }
    #endregion
}
