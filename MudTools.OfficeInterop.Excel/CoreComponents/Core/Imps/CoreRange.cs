//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;
using System.Reflection;

namespace MudTools.OfficeInterop.Excel.Imps;

internal abstract class CoreRange<T, TR> : ICoreRange<TR>
    where T : CoreRange<T, TR>, TR, new()
    where TR : ICoreRange<TR>
{

    #region 私有字段
    internal MsExcel.Range? _range; // 封装的Excel Range对象（COM对象）
    private bool _disposedValue;   // 资源释放标记
    /// <summary>
    /// 用于记录日志的静态日志记录器。
    /// </summary>
    private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

    // 格式相关缓存
    private IExcelFont? _font;
    private IExcelBorders? _borders;

    // 区域缓存
    private TR? _firstCell;
    private TR? _lastCell;
    private TR? _next;
    private TR? _previous;
    private TR? _currentRegion;
    private TR? _entireRow;
    private TR? _entireColumn;
    private TR? _usedRange;
    private TR? _parentRange;
    private IExcelCells? _cells;

    // 集合缓存
    private IExcelColumns? _columns;
    private IExcelHyperlinks? _hyperlinks;
    private IExcelSmartTags? _smartTags;
    private IExcelComment? _comment;
    private IExcelErrors? _errors = null;
    private IExcelPageSetup? _pageSetup;
    #endregion

    #region 属性

    /// <summary>
    /// 获取底层的Excel Range对象
    /// </summary>
    internal MsExcel.Range InternalRange
    {
        get => _range ?? throw new ObjectDisposedException(nameof(ExcelRange));
        private set => _range = value;
    }



    #region 基本属性
    public IExcelApplication Application => new ExcelApplication(_range.Application);

    /// <summary>
    /// 获取或设置单元格的值
    /// </summary>
    public object Value
    {
        get => InternalRange.Value;
        set => InternalRange.Value = value;
    }

    public object[,] ArrayValue
    {
        get
        {
            var obj = InternalRange.Value;
            if (obj is object[,] objs && objs != null)
                return objs;
            return null;
        }
        set => InternalRange.Value = value;
    }

    public int PageBreak
    {
        get => InternalRange.PageBreak;
        set => InternalRange.PageBreak = value;
    }

    public double? NumberValue
    {
        get
        {
            if (InternalRange.Value == null || InternalRange.Value == DBNull.Value)
                return null;
            return Convert.ToDouble(Value);
        }
        set
        {
            InternalRange.Value = value;
        }
    }

    public double[]? NumberValues
    {
        get
        {
            if (InternalRange.Value == null || InternalRange.Value == DBNull.Value)
                return null;
            if (InternalRange.Value is Array array)
            {
                List<double> values = [];
                foreach (var item in array)
                {
                    var dVla = Convert.ToDouble(item);
                    values.Add(dVla);
                }
            }
            return [];
        }
    }

    /// <summary>
    /// 获取或设置单元格的公式(A1格式)
    /// </summary>
    public object Formula
    {
        get => InternalRange.Formula;
        set => InternalRange.Formula = value;
    }

    public XlFormulaLabel FormulaLabel
    {
        get => InternalRange != null ? InternalRange.FormulaLabel.EnumConvert(XlFormulaLabel.xlNoLabels) : XlFormulaLabel.xlNoLabels;
        set
        {
            if (InternalRange != null)
                InternalRange.FormulaLabel = value.EnumConvert(MsExcel.XlFormulaLabel.xlNoLabels);
        }
    }


    public object PrefixCharacter
    {
        get => InternalRange.PrefixCharacter;
    }

    private IExcelFormatConditions? _excelFormatConditions;

    public IExcelFormatConditions? FormatConditions
    {
        get
        {
            if (_excelFormatConditions != null)
                return _excelFormatConditions;
            _excelFormatConditions = _range != null ? new ExcelFormatConditions(_range.FormatConditions) : null;
            return _excelFormatConditions;
        }
    }

    private IExcelCharacters? _characters;

    public IExcelCharacters? Characters
    {
        get
        {
            if (_characters != null)
                return _characters;
            _characters = _range != null ? new ExcelCharacters(_range.Characters) : null;
            return _characters;
        }
    }

    public bool HasFormula
    {
        get => InternalRange.HasFormula != null && Convert.ToBoolean(InternalRange.HasFormula);
    }

    /// <summary>
    /// 获取一个值，该值指示此区域是否是数组公式的一部分。
    /// </summary>
    public bool HasArray
    {
        get => InternalRange.HasArray != null && Convert.ToBoolean(InternalRange.HasArray);
    }

    /// <summary>
    /// 获取或设置数组公式
    /// </summary>
    public string FormulaArray
    {
        get => InternalRange.FormulaArray?.ToString();
        set => InternalRange.FormulaArray = value;
    }

    /// <summary>
    /// 获取或设置R1C1格式公式
    /// </summary>
    public object FormulaR1C1
    {
        get => InternalRange.FormulaR1C1;
        set => InternalRange.FormulaR1C1 = value;
    }

    /// <summary>
    /// 获取单元格的显示文本
    /// </summary>
    public string Text => InternalRange.Text?.ToString() ?? "";
    #endregion

    #region 尺寸与位置
    /// <summary>
    /// 获取区域的行数
    /// </summary>
    public int RowsCount => InternalRange.Rows.Count;

    /// <summary>
    /// 获取区域的列数
    /// </summary>
    public int ColumnsCount => InternalRange.Columns.Count;

    /// <summary>
    /// 获取区域的起始行号
    /// </summary>
    public int Row => InternalRange.Row;

    /// <summary>
    /// 获取区域的起始列号
    /// </summary>
    public int Column => InternalRange.Column;

    /// <summary>
    /// 获取区域左边缘的位置（以磅为单位）
    /// </summary>
    public double Left
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.Left); }
            catch { return 0; }
        }
    }

    /// <summary>
    /// 获取区域上边缘的位置（以磅为单位）
    /// </summary>
    public double Top
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.Top); }
            catch { return 0; }
        }
    }

    /// <summary>
    /// 获取区域的宽度（以磅为单位）
    /// </summary>
    public double Width
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.Width); }
            catch { return 0; }
        }
    }

    /// <summary>
    /// 获取区域的高度（以磅为单位）
    /// </summary>
    public double Height
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.Height); }
            catch { return 0; }
        }
    }

    /// <summary>
    /// 获取或设置行高（以磅为单位）
    /// </summary>
    public double RowHeight
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.RowHeight); }
            catch { return 0; }
        }
        set => InternalRange.RowHeight = value;
    }

    /// <summary>
    /// 获取或设置列宽（以字符数为单位）
    /// </summary>
    public double ColumnWidth
    {
        get
        {
            try { return Convert.ToDouble(InternalRange.ColumnWidth); }
            catch { return 0; }
        }
        set => InternalRange.ColumnWidth = value;
    }
    #endregion

    #region 格式与样式
    /// <summary>
    /// 获取字体格式对象
    /// </summary>
    public IExcelFont Font => _font ??= new ExcelFont(InternalRange.Font);

    /// <summary>
    /// 获取或设置背景色(RGB值)
    /// </summary>
    public int InteriorColor
    {
        get
        {
            try { return Convert.ToInt32(InternalRange.Interior.Color); }
            catch { return 0; }
        }
        set => InternalRange.Interior.Color = value;
    }

    /// <summary>
    /// 获取或设置水平对齐方式
    /// </summary>
    public XlHAlign HorizontalAlignment
    {
        get => InternalRange != null ? InternalRange.HorizontalAlignment.ObjectConvertEnum(XlHAlign.xlHAlignGeneral) : XlHAlign.xlHAlignGeneral;
        set
        {
            if (InternalRange != null)
                InternalRange.HorizontalAlignment = value.EnumConvert(MsExcel.XlHAlign.xlHAlignGeneral);
        }
    }

    /// <summary>
    /// 获取或设置垂直对齐方式
    /// </summary>
    public XlVAlign VerticalAlignment
    {
        get => InternalRange != null ? InternalRange.VerticalAlignment.ObjectConvertEnum(XlVAlign.xlVAlignBottom) : XlVAlign.xlVAlignBottom;
        set
        {
            if (InternalRange != null)
                InternalRange.VerticalAlignment = value.EnumConvert(MsExcel.XlVAlign.xlVAlignBottom);
        }
    }

    /// <summary>
    /// 获取或设置文本旋转角度（-90到90度）
    /// </summary>
    public XlOrientation Orientation
    {
        get => InternalRange != null ? InternalRange.Orientation.ObjectConvertEnum(XlOrientation.xlDownward) : XlOrientation.xlDownward;
        set
        {
            if (InternalRange != null)
                InternalRange.Orientation = value.EnumConvert(MsExcel.XlOrientation.xlDownward);
        }
    }

    /// <summary>
    /// 获取或设置是否自动换行
    /// </summary>
    public bool WrapText
    {
        get => Convert.ToBoolean(InternalRange.WrapText);
        set => InternalRange.WrapText = value;
    }

    public IExcelInterior? Interior
    {
        get
        {
            return InternalRange != null ? new ExcelInterior(InternalRange.Interior) : null;
        }
    }

    /// <summary>
    /// 获取或设置缩进级别
    /// </summary>
    public int IndentLevel
    {
        get => Convert.ToInt32(InternalRange.IndentLevel);
        set => InternalRange.IndentLevel = value;
    }

    /// <summary>
    /// 获取或设置单元格填充图案
    /// </summary>
    public XlPattern Pattern
    {
        get => InternalRange != null ? InternalRange.Interior.Pattern.ObjectConvertEnum(XlPattern.xlPatternNone) : XlPattern.xlPatternNone;
        set
        {
            if (InternalRange != null)
                InternalRange.Interior.Pattern = value.EnumConvert(MsExcel.XlPattern.xlPatternNone);
        }
    }

    /// <summary>
    /// 获取或设置单元格样式名称
    /// </summary>
    public IExcelStyle? Style
    {
        get => InternalRange != null ? new ExcelStyle(InternalRange.Style as MsExcel.Style) : null;
        set
        {
            if (value is ExcelStyle style)
                InternalRange.Style = style._style;
            else if (value is MsExcel.Style oStyle)
                InternalRange.Style = oStyle;
            else
                throw new NotSupportedException();
        }
    }

    /// <summary>
    /// 获取或设置图案颜色(RGB值)
    /// </summary>
    public Color PatternColor
    {
        get => ColorTranslator.FromOle(Convert.ToInt32(InternalRange.Interior.PatternColor));
        set => InternalRange.Interior.PatternColor = ColorTranslator.ToOle(value);
    }

    /// <summary>
    /// 获取或设置阅读顺序
    /// </summary>
    public int ReadingOrder
    {
        get => Convert.ToInt32(InternalRange.ReadingOrder);
        set => InternalRange.ReadingOrder = value;
    }
    public string NumberFormatLocal
    {
        get => _range.NumberFormatLocal?.ToString();
        set => NumberFormatLocal = value;
    }


    /// <summary>
    /// 获取或设置数字格式
    /// </summary>
    public string NumberFormat
    {
        get => InternalRange.NumberFormat?.ToString() ?? "";
        set => InternalRange.NumberFormat = value;
    }

    public IExcelPhonetics? Phonetics => InternalRange != null ? new ExcelPhonetics(InternalRange.Phonetics) : null;

    /// <summary>
    /// 获取单元格边框集合
    /// </summary>
    public IExcelBorders Borders => _borders ??= new ExcelBorders(InternalRange.Borders);

    public ExcelRectange RangeRect => new() { Left = Convert.ToInt32(Left), Top = Convert.ToInt32(Top), Height = Convert.ToInt32(Height), Width = Convert.ToInt32(Width) };
    #endregion

    #region 工作表与区域
    /// <summary>
    /// 获取所属工作表对象
    /// </summary>
    public IExcelWorksheet Worksheet => new ExcelWorksheet(InternalRange.Worksheet);

    /// <summary>
    /// 获取所属工作表名称
    /// </summary>
    public string WorksheetName => InternalRange.Worksheet?.Name ?? string.Empty;

    /// <summary>
    /// 获取区域的地址引用（A1样式）
    /// </summary>
    public string Address => InternalRange.Address;

    public IExcelRangeAddress RangeAddress => new ExcelRangeAddress(InternalRange);

    public string Name
    {
        get => InternalRange.Name?.ToString();
        set => InternalRange.Name = value;
    }

    /// <summary>
    /// 获取区域中的单元格数量
    /// </summary>
    public int Count => InternalRange.Count;

    /// <summary>
    /// 获取区域中的单元格数量
    /// </summary>
    public long CountLarge => Convert.ToInt64(InternalRange.CountLarge);

    /// <summary>
    /// 获取区域中的行集合
    /// </summary>
    public IExcelRows Rows => new ExcelRange(InternalRange.Rows);

    /// <summary>
    /// 获取区域中的列集合
    /// </summary>
    public IExcelColumns Columns => _columns ??= new ExcelRange(InternalRange.Columns);

    private IExcelAreas _areas;

    public IExcelAreas Areas => _areas ??= new ExcelAreas(InternalRange.Areas);

    /// <summary>
    /// 获取区域中的单元格对象。
    /// </summary>
    public IExcelCells Cells => _cells ??= new ExcelCells(InternalRange.Cells);

    /// <summary>
    /// 获取当前区域的第一个单元格
    /// </summary>
    public TR FirstCell => GetProperty((MsExcel.Range)InternalRange.Cells[1, 1], ref _firstCell);

    /// <summary>
    /// 获取当前区域的最后一个单元格
    /// </summary>
    public TR LastCell => GetProperty((MsExcel.Range)InternalRange.Cells[RowsCount, ColumnsCount], ref _lastCell);

    /// <summary>
    /// 获取下一个相邻区域
    /// </summary>
    public TR Next => GetProperty(InternalRange.Next, ref _next);

    /// <summary>
    /// 获取上一个相邻区域
    /// </summary>
    public TR Previous => GetProperty(InternalRange.Previous, ref _previous);

    /// <summary>
    /// 获取当前区域（包含连续数据的区域）
    /// </summary>
    public TR CurrentRegion => GetProperty(InternalRange.CurrentRegion, ref _currentRegion);

    /// <summary>
    /// 获取整行区域
    /// </summary>
    public TR EntireRow => GetProperty(InternalRange.EntireRow, ref _entireRow);

    /// <summary>
    /// 获取整列区域
    /// </summary>
    public TR EntireColumn => GetProperty(InternalRange.EntireColumn, ref _entireColumn);

    /// <summary>
    /// 获取工作表的已使用区域
    /// </summary>
    public TR UsedRange => GetProperty(InternalRange.Worksheet.UsedRange, ref _usedRange);

    /// <summary>
    /// 获取父区域
    /// </summary>
    public TR ParentRange => GetProperty(InternalRange.Parent as MsExcel.Range, ref _parentRange);

    /// <summary>
    /// 获取父对象（可能是工作表或区域）
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_range?.Parent == null)
            {
                return null;
            }
            if (_range.Parent is MsExcel.Worksheet worksheet)
            {
                return ParentSheet;
            }
            return _range.Parent;
        }
    }

    private IExcelWorksheet _parentSheet;

    public IExcelWorksheet? ParentSheet
    {
        get
        {
            if (_parentSheet != null)
                return _parentSheet;

            if (_range?.Parent == null)
            {
                return null;
            }
            if (_range.Parent is MsExcel.Worksheet worksheet)
            {
                _parentSheet ??= new ExcelWorksheet(worksheet);
            }
            return _parentSheet;
        }
    }
    #endregion

    #region 超链接与注释
    /// <summary>
    /// 获取或设置超链接地址
    /// </summary>
    public string? Hyperlink
    {
        get
        {
            try
            {
                return InternalRange.Hyperlinks.Count > 0
                    ? InternalRange.Hyperlinks[1].Address
                    : null;
            }
            catch { return null; }
        }
        set
        {
            try
            {
                // 清除现有超链接
                InternalRange.Hyperlinks.Delete();

                // 添加新超链接
                if (!string.IsNullOrEmpty(value))
                    InternalRange.Hyperlinks.Add(InternalRange, value);
            }
            catch { }
        }
    }

    /// <summary>
    /// 获取超链接集合
    /// </summary>
    public IExcelHyperlinks Hyperlinks => _hyperlinks ??= new ExcelHyperlinks(InternalRange.Hyperlinks);

    /// <summary>
    /// 获取单元格批注对象
    /// </summary>
    public IExcelComment Comment => _comment ??= new ExcelComment(InternalRange.Comment);

    public IExcelErrors Errors => _errors ?? new ExcelErrors(InternalRange.Errors);

    /// <summary>
    /// 获取或设置批注文本
    /// </summary>
    public string? CommentText
    {
        get => InternalRange.Comment?.Text();
        set
        {
            if (InternalRange.Comment != null)
                InternalRange.Comment.Text(value ?? string.Empty);
            else if (!string.IsNullOrEmpty(value))
                AddComment(value);
        }
    }
    #endregion

    #region 保护与打印
    /// <summary>
    /// 获取或设置是否隐藏行/列
    /// </summary>
    public bool Hidden
    {
        get
        {
            try { return Convert.ToBoolean(InternalRange.Hidden); }
            catch { return false; }
        }
        set => InternalRange.Hidden = value;
    }

    /// <summary>
    /// 获取区域是否包含合并单元格
    /// </summary>
    public bool MergeCells
    {
        get
        {
            try { return Convert.ToBoolean(InternalRange.MergeCells); }
            catch { return false; }
        }
    }

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// 对应 Range.Range 属性
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    public TR? Range(object? cell1, object? cell2 = null)
    {
        try
        {
            if (cell1 is CoreRange<T, TR> range1)
                cell1 = range1.InternalRange;
            if (cell2 is CoreRange<T, TR> range2)
                cell2 = range2.InternalRange;

            cell1 ??= Type.Missing;
            cell2 ??= Type.Missing;

            var range = _range?.Range[cell1, cell2];
            return CreateRangeObject(range);
        }
        catch
        {
            return default;
        }
    }



    /// <summary>
    /// 获取或设置是否打印网格线
    /// </summary>
    public bool PrintGridlines
    {
        get => InternalRange.Worksheet.PageSetup.PrintGridlines;
        set => InternalRange.Worksheet.PageSetup.PrintGridlines = value;
    }

    /// <summary>
    /// 获取或设置是否打印行号列标
    /// </summary>
    public bool PrintHeadings
    {
        get => InternalRange.Worksheet.PageSetup.PrintHeadings;
        set => InternalRange.Worksheet.PageSetup.PrintHeadings = value;
    }

    /// <summary>
    /// 获取页面设置对象
    /// </summary>
    public IExcelPageSetup PageSetup => _pageSetup ??= new ExcelPageSetup(InternalRange.Worksheet.PageSetup);

    /// <summary>
    /// 获取或设置单元格是否锁定
    /// </summary>
    public bool Locked
    {
        get => Convert.ToBoolean(InternalRange.Locked);
        set => InternalRange.Locked = value;
    }

    /// <summary>
    /// 获取或设置公式是否隐藏
    /// </summary>
    public bool FormulaHidden
    {
        get => Convert.ToBoolean(InternalRange.FormulaHidden);
        set => InternalRange.FormulaHidden = value;
    }
    #endregion

    #region 智能标记
    /// <summary>
    /// 获取智能标记集合
    /// </summary>
    public IExcelSmartTags SmartTags => _smartTags ??= new ExcelSmartTags(InternalRange.SmartTags);
    #endregion

    protected TR? GetProperty(MsExcel.Range rang, ref TR? rObj)
    {
        if (rObj != null)
            return rObj;
        rObj = CreateRangeObject(rang);
        return rObj;
    }

    protected TR? CreateRangeObject(MsExcel.Range rang)
    {
        if (rang == null) return default;

        var rObj = new T() { InternalRange = rang };
        return rObj;
    }

    #region 新增属性
    private TR? _mergeArea;

    /// <summary>
    /// 获取包含指定单元格的合并区域
    /// </summary>
    /// <remarks>
    /// 若单元格不属于合并区域，则返回单元格自身
    /// </remarks>
    public TR MergeArea => GetProperty(InternalRange.MergeArea, ref _mergeArea);

    /// <summary>
    /// 获取或设置数据验证规则对象
    /// </summary>
    public IExcelValidation Validation
    {
        get
        {
            try { return new ExcelValidation(InternalRange.Validation, _range); }
            catch { throw new InvalidOperationException("数据验证功能不可用"); }
        }
    }

    /// <summary>
    /// 获取或设置分级显示展开状态
    /// </summary>
    public bool ShowDetail
    {
        get => Convert.ToBoolean(InternalRange.ShowDetail);
        set => InternalRange.ShowDetail = value;
    }
    #endregion
    #endregion

    #region 方法
    #region 构造函数
    /// <summary>
    /// 初始化Excel范围对象
    /// </summary>
    /// <param name="range">Excel原生Range对象</param>
    /// <exception cref="ArgumentNullException">range参数为空时抛出</exception>
    public CoreRange(MsExcel.Range? range)
    {
        _range = range;
        _disposedValue = false;
    }
    #endregion

    #region 新增方法
    /// <summary>
    /// 激活当前区域
    /// </summary>
    public void Activate()
    {
        try { InternalRange.Activate(); }
        catch (COMException ex)
        {
            throw new InvalidOperationException("激活区域失败", ex);
        }
    }

    /// <summary>
    /// 计算公式结果
    /// </summary>
    /// <remarks>
    /// 强制重新计算区域内的所有公式
    /// </remarks>
    public void Calculate()
    {
        try { InternalRange.Calculate(); }
        catch (COMException ex)
        {
            throw new InvalidOperationException("公式计算失败", ex);
        }
    }

    /// <summary>
    /// 自动填充数据到目标区域
    /// </summary>
    /// <param name="destination">填充目标区域</param>
    /// <param name="type">填充类型</param>
    public void AutoFill(TR destination, AutoFillType type = AutoFillType.xlFillDefault)
    {
        if (destination is CoreRange<T, TR> excelRange)
        {
            InternalRange.AutoFill(
                Destination: excelRange.InternalRange,
                Type: (MsExcel.XlAutoFillType)(int)type
            );
        }
        else
        {
            throw new ExcelOperationException("目标区域必须是ExcelRange类型");
        }
    }

    /// <summary>
    /// 获取区域的本地化地址。
    /// </summary>
    /// <param name="rowAbsolute">指定行号是否为绝对引用 (例如，$1)。默认为 true。</param>
    /// <param name="columnAbsolute">指定列号是否为绝对引用 (例如，$A)。默认为 true。</param>
    /// <param name="referenceStyle">指定地址样式。默认为 xlA1。</param>
    /// <param name="external">如果为 true，则返回包含工作簿和工作表名称的外部引用。默认为 false。</param>
    /// <param name="relativeTo">如果 ReferenceStyle 为 xlR1C1，则指定相对引用的起始点。</param>
    /// <returns>表示区域地址的字符串。</returns>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败，可能会抛出 COM 异常。
    /// </exception>
    /// <exception cref="ArgumentNullException">
    /// 如果内部的 _range 对象为 null。
    /// </exception>
    public string GetAddressLocal(
        bool? rowAbsolute = true,
        bool? columnAbsolute = true,
        XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
        bool? external = false,
        object? relativeTo = null)
    {
        // 检查内部对象是否为 null
        if (_range == null)
        {
            throw new ArgumentNullException(nameof(_range), "Underlying Range object is null.");
        }

        try
        {
            return _range.get_AddressLocal(
                rowAbsolute.ComArgsVal(),
                columnAbsolute.ComArgsVal(),
                (MsExcel.XlReferenceStyle)referenceStyle,
                external.ComArgsVal(),
                relativeTo ?? Type.Missing
            );
        }
        catch (COMException comEx)
        {
            _log.Error("COM Exception in GetAddressLocal", comEx);
            throw new ExcelOperationException("COM Exception in GetAddressLocal", comEx);
        }
        catch (Exception ex)
        {
            _log.Error("General Exception in GetAddressLocal", ex);
            throw new InvalidOperationException("Failed to get AddressLocal.", ex);
        }
    }


    /// <summary>
    /// 获取区域的本地化地址，使用默认参数 (A1, 绝对引用, 非外部)。
    /// </summary>
    public string AddressLocal
    {
        get
        {
            return GetAddressLocal();
        }
    }


    /// <summary>
    /// 直接替换原始调用：range.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing)
    /// </summary>
    /// <returns>地址字符串</returns>
    public string? GetDefaultA1Address()
    {
        return GetAddress(null, null, XlReferenceStyle.xlA1, null, null);
    }


    /// <summary>
    /// 完全兼容原始调用的静态方法
    /// </summary>
    /// <param name="rowAbsolute">行是否绝对引用</param>
    /// <param name="columnAbsolute">列是否绝对引用</param>
    /// <param name="referenceStyle">引用样式</param>
    /// <param name="external">是否外部引用</param>
    /// <param name="relativeTo">相对引用基准</param>
    /// <returns>地址字符串</returns>
    public string? GetAddress(
        bool? rowAbsolute = true,
        bool? columnAbsolute = true,
        XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
        bool? external = false,
        object? relativeTo = null)
    {
        try
        {
            return _range?.get_Address(
                            rowAbsolute.ComArgsVal(),
                            columnAbsolute.ComArgsVal(),
                            (MsExcel.XlReferenceStyle)(int)referenceStyle,
                            external.ComArgsVal(),
                            relativeTo ?? Type.Missing
                            );
        }
        catch (COMException comEx)
        {
            _log.Error("COM Exception in GetAddress", comEx);
            throw new ExcelOperationException("COM Exception in GetAddress", comEx);
        }
        catch (Exception ex)
        {
            _log.Error("General Exception in GetAddress", ex);
            throw new InvalidOperationException("Failed to get GetAddress.", ex);
        }
    }

    /// <summary>
    /// 获取区域内的直接依赖单元格
    /// </summary>
    /// <returns>直接前驱单元格区域</returns>
    public TR? GetDirectDependents()
    {
        var dependents = InternalRange.DirectDependents;
        return CreateRangeObject(dependents);
    }

    /// <summary>
    /// 获取区域内的直接引用单元格
    /// </summary>
    /// <returns>直接引用单元格区域</returns>
    public TR? GetDirectPrecedents()
    {
        var precedents = InternalRange.DirectPrecedents;
        return CreateRangeObject(precedents);
    }
    #endregion

    #region 编辑操作
    /// <summary>
    /// 复制当前区域到剪贴板
    /// </summary>
    /// <returns>操作是否成功</returns>
    public bool Copy()
    {
        try
        {
            InternalRange.Copy();
            return true;
        }
        catch { return false; }
    }
    public void CopyPicture(XlPictureAppearance appearance = XlPictureAppearance.xlScreen, XlCopyPictureFormat format = XlCopyPictureFormat.xlPicture)
    {
        if (InternalRange == null) return;

        try
        {
            InternalRange.CopyPicture(
                appearance.EnumConvert(MsExcel.XlPictureAppearance.xlScreen),
                format.EnumConvert(MsExcel.XlCopyPictureFormat.xlPicture));
        }
        catch (Exception ex)
        {
            _log.Error($"复制图片操作失败: {ex.Message}", ex);
            throw;
        }

    }

    /// <summary>
    /// 复制当前区域到目标区域
    /// </summary>
    /// <param name="destination">目标区域</param>
    /// <exception cref="ArgumentException">目标区域类型无效时抛出</exception>
    public void Copy(TR destination)
    {
        if (destination is CoreRange<T, TR> excelRange)
        {
            InternalRange.Copy(excelRange.InternalRange);
        }
        else
        {
            throw new ExcelOperationException("目标区域必须是ExcelRange类型");
        }
    }

    /// <summary>
    /// 复制Range区域并粘贴到指定位置
    /// </summary>
    /// <param name="targetAddress">目标地址</param>
    /// <param name="pasteType">粘贴类型</param>
    /// <returns>是否操作成功</returns>
    public bool CopyAndPaste(string targetAddress, XlPasteType pasteType = XlPasteType.xlPasteAll)
    {
        try
        {
            // 复制源区域
            _range?.Copy();

            // 获取目标区域
            MsExcel.Range? targetRange = _range?.Worksheet.Range[targetAddress];

            // 粘贴到目标位置
            targetRange?.PasteSpecial((MsExcel.XlPasteType)pasteType);

            // 清除剪贴板
            this.Application.CutCopyMode = XlCutCopyMode.xlCopy;

            return true;
        }
        catch (Exception ex)
        {
            _log.Error($"复制粘贴操作失败: {ex.Message}", ex);
            return false;
        }
    }


    /// <summary>
    /// 粘贴数据到当前区域
    /// </summary>
    /// <param name="from">源区域</param>
    /// <param name="type">粘贴类型</param>
    /// <param name="operation">粘贴操作</param>
    /// <param name="skipBlanks">是否跳过空单元格</param>
    /// <param name="transpose">是否转置</param>
    /// <returns>操作是否成功</returns>
    public bool CopyAndPaste(TR from,
        PasteType type = PasteType.All,
        PasteOperation operation = PasteOperation.None,
        bool skipBlanks = false,
        bool transpose = false)
    {
        if (from == null) return false;
        try
        {
            if (!from.Copy())
                return false;

            InternalRange.PasteSpecial(
                Paste: (MsExcel.XlPasteType)type,
                Operation: (MsExcel.XlPasteSpecialOperation)operation,
                SkipBlanks: skipBlanks,
                Transpose: transpose);

            return true;
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"粘贴数据到当前区域失败: {ex.Message}", ex);
        }
    }

    public TR? PasteSpecial(
        XlPasteType paste = XlPasteType.xlPasteAll,
        XlPasteSpecialOperation operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone,
        bool? skipBlanks = false,
        bool? transpose = false)
    {

        object? obj = _range?.PasteSpecial((MsExcel.XlPasteType)(int)paste,
                            (MsExcel.XlPasteSpecialOperation)(int)operation,
                            skipBlanks.ComArgsVal(),
                            transpose.ComArgsVal());
        if (obj is MsExcel.Range rang)
        {
            return CreateRangeObject(rang);
        }
        return default;
    }


    /// <summary>
    /// 插入单元格
    /// </summary>
    /// <param name="direction">移动方向</param>
    /// <param name="origin">格式来源</param>
    /// <returns>操作是否成功</returns>
    public bool Insert(
        XlDirection direction = XlDirection.xlDown,
        XlInsertFormatOrigin origin = XlInsertFormatOrigin.FromRightOrBelow)
    {
        try
        {
            InternalRange?.Insert(
                Shift: (MsExcel.XlInsertShiftDirection)direction,
                CopyOrigin: (MsExcel.XlInsertFormatOrigin)origin);
            return true;
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"插入单元格失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 删除单元格
    /// </summary>
    /// <param name="direction">移动方向</param>
    /// <returns>操作是否成功</returns>
    public bool Delete(XlDirection direction = XlDirection.xlToLeft)
    {
        try
        {
            InternalRange?.Delete(Shift: (MsExcel.XlDeleteShiftDirection)direction);
            return true;
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"删除单元格失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 清除单元格内容
    /// </summary>
    public void ClearContents()
    {
        InternalRange?.ClearContents();
    }

    /// <summary>
    /// 清除所有内容（值、格式等）
    /// </summary>
    public void Clear()
    {
        InternalRange?.Clear();
    }

    public object Parse(string? parseLine = null, TR? range = default)
    {
        object? r = Type.Missing;
        if (range != null && range is CoreRange<T, TR> otherRange)
            r = otherRange.InternalRange;

        return InternalRange?.Parse(parseLine.ComArgsVal(), r);
    }

    /// <summary>
    /// 清除格式设置
    /// </summary>
    public void ClearFormats()
    {
        InternalRange.ClearFormats();
    }
    #endregion

    #region 区域操作

    public object? BorderAround(XlLineStyle? lineStyle = null,
         XlBorderWeight weight = XlBorderWeight.xlThin,
         XlColorIndex colorIndex = XlColorIndex.xlColorIndexAutomatic,
         Color? color = null)
    {
        object lineStyleObj = Type.Missing;
        if (lineStyle != null)
            lineStyleObj = (MsExcel.XlLineStyle)(int)lineStyle;
        object colorObj = Type.Missing;
        if (color != null)
            colorObj = color.Value.ToArgb();

        return _range?.BorderAround(lineStyleObj,
            (MsExcel.XlBorderWeight)(int)weight,
            (MsExcel.XlColorIndex)(int)colorIndex,
            colorObj);
    }
    /// <summary>
    /// 合并单元格
    /// </summary>
    /// <param name="merge">是否合并（默认为true）</param>
    public void Merge(bool merge = true)
    {
        InternalRange.Merge(merge);
    }

    /// <summary>
    /// 取消单元格合并
    /// </summary>
    public void UnMerge()
    {
        InternalRange.UnMerge();
    }

    /// <summary>
    /// 调整区域大小
    /// </summary>
    /// <param name="rowSize">新行数（若为-1则保持原行数）</param>
    /// <param name="columnSize">新列数（若为-1则保持原列数）</param>
    /// <returns>调整后的新区域</returns>
    public TR? Resize(int? rowSize = -1, int? columnSize = -1)
    {
        rowSize = rowSize ?? -1;
        columnSize = columnSize ?? -1;

        // 保持原尺寸的逻辑处理
        rowSize = rowSize < 0 ? RowsCount : rowSize;
        columnSize = columnSize < 0 ? ColumnsCount : columnSize;

        var range = InternalRange.Resize[rowSize, columnSize];
        return CreateRangeObject(range);
    }

    /// <summary>
    /// 获取偏移后的区域
    /// </summary>
    /// <param name="rowOffset">行偏移量</param>
    /// <param name="columnOffset">列偏移量</param>
    /// <returns>新区域对象</returns>
    public TR? Offset(int? rowOffset = 0, int? columnOffset = 0)
    {
        rowOffset = rowOffset ?? 0;
        columnOffset = columnOffset ?? 0;

        return CreateRangeObject(InternalRange.Offset[rowOffset, columnOffset]);
    }

    public TR? Offset(long? rowOffset = 0, long? columnOffset = 0)
    {
        rowOffset = rowOffset ?? 0;
        columnOffset = columnOffset ?? 0;
        return CreateRangeObject(InternalRange.Offset[rowOffset, columnOffset]);
    }

    /// <summary>
    /// 获取与另一区域的交集区域
    /// </summary>
    /// <param name="other">另一区域</param>
    /// <returns>交集区域（无交集时返回null）</returns>
    public TR? Intersect(TR other)
    {
        if (other is not CoreRange<T, TR> otherRange)
            return default;

        try
        {
            MsExcel.Range intersect = InternalRange.Application.Intersect(InternalRange, otherRange.InternalRange);
            return CreateRangeObject(intersect);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"获取与另一区域的交集区域失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 获取与另一区域的并集区域
    /// </summary>
    /// <param name="other">另一区域</param>
    /// <returns>并集区域</returns>
    public TR? Union(TR other)
    {
        if (other is not CoreRange<T, TR> otherRange)
            return default;
        try
        {
            MsExcel.Range union = InternalRange.Application.Union(InternalRange, otherRange.InternalRange);

            return CreateRangeObject(union);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"获取与另一区域的并集区域失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 自动调整列宽
    /// </summary>
    public void AutoFit()
    {
        InternalRange?.AutoFit();
    }

    public object? AutoFormat(XlRangeAutoFormat format = XlRangeAutoFormat.xlRangeAutoFormatClassic1,
        bool? number = true, bool? font = true, bool? alignment = true,
        bool? border = true, bool? pattern = true, bool? width = true)
    {
        return InternalRange?.AutoFormat(
            Format: (MsExcel.XlRangeAutoFormat)(int)format,
            Number: number, Font: font, Alignment: alignment,
            Border: border, Pattern: pattern, Width: width);
    }

    public object? AutoOutline()
    {
        return InternalRange?.AutoOutline();
    }


    /// <summary>
    /// 获取区域边缘的单元格
    /// </summary>
    /// <param name="direction">方向</param>
    /// <returns>目标单元格</returns>
    public TR? End(XlDirection direction = XlDirection.xlDown)
    {
        try
        {
            return CreateRangeObject(InternalRange.End[(MsExcel.XlDirection)direction]);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"获取区域边缘的单元格失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 选中当前区域
    /// </summary>
    /// <returns>操作是否成功</returns>
    public bool Select()
    {
        try
        {
            InternalRange.Select();
            return true;
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException($"选中当前区域失败: {ex.Message}", ex);
        }
    }
    #endregion

    #region 筛选与排序
    /// <summary>
    /// 应用自动筛选
    /// </summary>
    public void AutoFilter()
    {
        InternalRange.AutoFilter();
    }

    /// <summary>
    /// 移除自动筛选
    /// </summary>
    public void RemoveAutoFilter()
    {
        if (InternalRange.Worksheet.AutoFilter != null &&
            InternalRange.Worksheet.AutoFilter.Range != null &&
            InternalRange.Worksheet.AutoFilter.Range.Address == InternalRange.Address)
        {
            InternalRange.Worksheet.AutoFilterMode = false;
        }
    }

    /// <summary>
    /// 对区域进行排序
    /// </summary>
    /// <param name="key1">主要排序键</param>
    /// <param name="key2">次要排序键</param>
    /// <param name="type">排序类型</param>
    /// <param name="key3">第三排序键</param>
    /// <param name="orderCustom">自定义排序顺序</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="order1">主要排序顺序</param>
    /// <param name="order2">次要排序顺序</param>
    /// <param name="order3">第三排序顺序</param>
    /// <param name="header">是否包含标题行</param>
    /// <param name="orientation">排序方向</param>
    /// <param name="sortMethod">排序方法</param>
    /// <param name="dataOption1">主要键数据选项</param>
    /// <param name="dataOption2">次要键数据选项</param>
    /// <param name="dataOption3">第三键数据选项</param>
    public void Sort(
        object key1,
        object key2,
        object type,
        object key3,
        object orderCustom,
        object matchCase,
        XlSortOrder order1 = XlSortOrder.xlAscending,
        XlSortOrder order2 = XlSortOrder.xlAscending,
        XlSortOrder order3 = XlSortOrder.xlAscending,
        XlYesNoGuess header = XlYesNoGuess.xlNo,
        XlSortOrientation orientation = XlSortOrientation.xlSortRows,
        XlSortMethod sortMethod = XlSortMethod.xlPinYin,
        XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
        XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal,
        XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal)
    {
        InternalRange.Sort(
            Key1: key1,
            Order1: order1.EnumConvert(MsExcel.XlSortOrder.xlAscending),
            Key2: key2,
            Order2: order2.EnumConvert(MsExcel.XlSortOrder.xlAscending),
            Key3: key3,
            Order3: order3.EnumConvert(MsExcel.XlSortOrder.xlAscending),
            Header: header.EnumConvert(MsExcel.XlYesNoGuess.xlNo),
            OrderCustom: orderCustom,
            MatchCase: matchCase,
            Orientation: orientation.EnumConvert(MsExcel.XlSortOrientation.xlSortRows),
            SortMethod: sortMethod.EnumConvert(MsExcel.XlSortMethod.xlPinYin),
            DataOption1: dataOption1.EnumConvert(MsExcel.XlSortDataOption.xlSortNormal),
            DataOption2: dataOption2.EnumConvert(MsExcel.XlSortDataOption.xlSortNormal),
            DataOption3: dataOption3.EnumConvert(MsExcel.XlSortDataOption.xlSortNormal));
    }

    #endregion

    #region 查找与替换
    public TR? FindNext(TR? after = default)
    {
        var afterObj = Type.Missing;
        if (after != null && after is CoreRange<T, TR> excelRange)
            afterObj = excelRange.InternalRange;
        var range = InternalRange.FindNext(afterObj);
        return CreateRangeObject(range);
    }

    public TR? FindPrevious(TR? after = default)
    {
        var afterObj = Type.Missing;
        if (after != null && after is CoreRange<T, TR> excelRange)
            afterObj = excelRange.InternalRange;
        var range = InternalRange.FindNext(afterObj);
        return CreateRangeObject(range);
    }


    /// <summary>
    /// 查找内容
    /// </summary>
    /// <param name="what">要查找的内容</param>
    /// <param name="after">开始搜索的位置</param>
    /// <param name="lookIn">搜索范围</param>
    /// <param name="lookAt">匹配方式</param>
    /// <param name="searchOrder">搜索顺序</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchByte">双字节匹配</param>
    /// <param name="searchFormat">搜索格式</param>
    /// <param name="searchDirection">搜索方向</param>
    /// <returns>找到的单元格</returns>
    public TR? Find(object what,
       TR? after = default,
        XlFindLookIn? lookIn = null,
        XlLookAt? lookAt = null,
        XlSearchOrder? searchOrder = null,
        XlSearchDirection searchDirection = XlSearchDirection.xlNext,
        bool? matchCase = null,
        bool? matchByte = null,
        object? searchFormat = null)
    {
        var afterObj = Type.Missing;
        if (after != null && after is CoreRange<T, TR> excelRange)
            afterObj = excelRange.InternalRange;


        object lookInObj = Type.Missing;
        if (lookIn != null)
            lookInObj = lookIn.EnumConvert(MsExcel.XlFindLookIn.xlValues);

        object lookAtObj = Type.Missing;
        if (lookAt != null)
            lookAtObj = lookAt.EnumConvert(MsExcel.XlLookAt.xlPart);

        object searchOrderObj = Type.Missing;
        if (searchOrder != null)
            searchOrderObj = searchOrder.EnumConvert(MsExcel.XlSearchOrder.xlByRows);

        searchFormat ??= Type.Missing;

        MsExcel.Range range = InternalRange.Find(
                    What: what,
                    After: afterObj,
                    LookIn: lookInObj,
                    LookAt: lookAtObj,
                    SearchOrder: searchOrderObj,
                    SearchDirection: searchDirection.EnumConvert(MsExcel.XlSearchDirection.xlNext),
                    MatchCase: matchCase.ComArgsVal(),
                    MatchByte: matchByte.ComArgsVal(),
                    SearchFormat: searchFormat);

        return CreateRangeObject(range);
    }

    public bool Replace(object What, object Replacement,
        XlLookAt? LookAt = null, XlSearchOrder? SearchOrder = null,
        bool? MatchCase = null, bool? MatchByte = null,
        object? SearchFormat = null, object? ReplaceFormat = null)
    {
        object lookAtObj = Type.Missing;
        if (LookAt != null)
            lookAtObj = LookAt.EnumConvert(MsExcel.XlLookAt.xlPart);

        object searchOrderObj = Type.Missing;
        if (SearchOrder != null)
            searchOrderObj = SearchOrder.EnumConvert(MsExcel.XlSearchOrder.xlByRows);

        SearchFormat ??= Type.Missing;
        ReplaceFormat ??= Type.Missing;

        return InternalRange.Replace(What: What, Replacement: Replacement,
                        LookAt: lookAtObj, SearchOrder: searchOrderObj,
                        MatchCase: MatchCase.ComArgsVal(), MatchByte: MatchByte.ComArgsVal(),
                        SearchFormat: SearchFormat,
                        ReplaceFormat: ReplaceFormat);
    }


    /// <summary>
    /// 获取特殊单元格
    /// </summary>
    /// <param name="type">单元格类型</param>
    /// <param name="value">特殊值</param>
    /// <returns>特殊单元格区域</returns>
    public TR? SpecialCells(XlCellType type, object? value = null)
    {
        try
        {
            value ??= Type.Missing;
            MsExcel.Range range = InternalRange.SpecialCells(type.EnumConvert(MsExcel.XlCellType.xlCellTypeBlanks), value);
            return CreateRangeObject(range);
        }
        catch
        {
            return default;
        }
    }


    #endregion

    #region 注释操作
    /// <summary>
    /// 添加批注
    /// </summary>
    /// <param name="text">批注文本</param>
    /// <returns>批注对象</returns>
    public IExcelComment? AddComment(string? text)
    {
        MsExcel.Comment? comment = InternalRange?.AddComment(text.ComArgsVal());
        return comment != null ? new ExcelComment(comment) : null;
    }

    /// <summary>
    /// 删除批注
    /// </summary>
    public void DeleteComment()
    {
        InternalRange.Comment?.Delete();
    }

    /// <summary>
    /// 清除所有批注
    /// </summary>
    public void ClearComments() => InternalRange.ClearComments();
    #endregion

    #region 超链接操作
    /// <summary>
    /// 添加超链接
    /// </summary>
    /// <param name="address">链接地址</param>
    /// <param name="subAddress">子地址（如工作表引用）</param>
    /// <param name="screenTip">屏幕提示文本</param>
    /// <param name="textToDisplay">显示文本</param>
    /// <returns>超链接对象</returns>
    public IExcelHyperlink AddHyperlink(string address, string? subAddress = null, string? screenTip = null, string? textToDisplay = null)
    {
        object hyperlink = InternalRange.Hyperlinks.Add(
            Anchor: InternalRange,
            Address: address,
            SubAddress: subAddress,
            ScreenTip: screenTip,
            TextToDisplay: textToDisplay);

        return new ExcelHyperlink((MsExcel.Hyperlink)hyperlink);
    }

    /// <summary>
    /// 清除所有超链接
    /// </summary>
    public void ClearHyperlinks() => InternalRange.Hyperlinks.Delete();
    #endregion

    #region 资源释放
    /// <summary>
    /// 释放托管和非托管资源
    /// </summary>
    /// <param name="disposing">是否主动释放</param>
    protected void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源
            DisposeManagedResources();
        }

        // 释放非托管资源(COM对象)
        ReleaseUnmanagedResources();

        _disposedValue = true;
    }

    /// <summary>
    /// 释放所有托管资源
    /// </summary>
    private void DisposeManagedResources()
    {
        // 统一释放所有实现了 IDisposable 的字段并置 null
        var fields = new IDisposable?[]
        {
        _font, _borders,_mergeArea,_characters,
        _firstCell, _lastCell, _next, _previous,
        _currentRegion, _entireRow, _entireColumn,
        _usedRange, _parentRange,_parentSheet,_errors,
        _hyperlinks, _smartTags, _comment, _pageSetup,
        _columns,_cells,_excelFormatConditions
        };

        foreach (var field in fields)
            field?.Dispose();

        // 所有字段置 null
        _characters = null;
        _parentSheet = null;
        _mergeArea = default;
        _font = null;
        _borders = null;
        _firstCell = default;
        _lastCell = default;
        _next = default;
        _previous = default;
        _currentRegion = default;
        _entireRow = default;
        _entireColumn = default;
        _usedRange = default;
        _parentRange = default;
        _columns = null;
        _hyperlinks = null;
        _smartTags = null;
        _comment = null;
        _errors = null;
        _pageSetup = null;
        _excelFormatConditions = null;
    }

    /// <summary>
    /// 释放非托管资源(COM对象)
    /// </summary>
    private void ReleaseUnmanagedResources()
    {
        if (_range != null)
        {
            Marshal.ReleaseComObject(_range);
            _range = null;
        }
    }

    ~CoreRange()
    {
        Dispose(false);
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    public IEnumerator<TR> GetEnumerator()
    {
        for (int iRow = 1; iRow <= _range.Rows.Count; iRow++)
        {
            for (int iCol = 1; iCol <= _range.Columns.Count; iCol++)
            {
                var range = _range[iRow, iCol] as MsExcel.Range;
                var t = CreateRangeObject(range);
                yield return t;

            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion
    #endregion
}
