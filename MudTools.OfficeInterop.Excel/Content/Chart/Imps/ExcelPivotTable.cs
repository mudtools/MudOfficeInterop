//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PivotTable 对象的二次封装实现类
/// 实现 IExcelPivotTable 接口
/// </summary>
internal class ExcelPivotTable : IExcelPivotTable
{
    internal MsExcel.PivotTable? _pivotTable;
    private bool _disposedValue = false;

    internal ExcelPivotTable(MsExcel.PivotTable pivotTable)
    {
        _pivotTable = pivotTable ?? throw new ArgumentNullException(nameof(pivotTable));
    }

    #region 基础属性
    public string Name
    {
        get => _pivotTable?.Name ?? string.Empty;
        set
        {
            if (_pivotTable != null)
            {
                _pivotTable.Name = value;
            }
        }
    }

    public object? Parent => _pivotTable?.Parent;

    public IExcelApplication? Application => _pivotTable != null ? new ExcelApplication(_pivotTable.Application) : null;

    public IExcelPivotCache? PivotCache() => _pivotTable != null ? new ExcelPivotCache(_pivotTable.PivotCache()) : null;

    public object? SourceData
    {
        get
        {
            var srcObj = CreateSourceObj(_pivotTable?.SourceData);
            return srcObj;
        }
        set
        {
            if (_pivotTable != null && value != null)
            {
                object? val = GetSourceObj(value);
                if (val != null) _pivotTable.SourceData = val;
            }
        }
    }

    private object? GetSourceObj(object? sourceData)
    {
        object? comSourceData = null;
        if (sourceData is ExcelRange rrange && rrange.InternalRange != null)
            comSourceData = rrange.InternalRange;
        else if (sourceData is ExcelListObject lo && lo.InternalComObject != null)
            comSourceData = lo.InternalComObject;
        else if (sourceData is ExcelPivotTable dt && dt._pivotTable != null)
            comSourceData = dt._pivotTable;
        else if (sourceData is string sourceString)
            comSourceData = sourceString;
        else
            comSourceData = null;
        return comSourceData;
    }

    private object CreateSourceObj(object? sourceData)
    {
        object comSourceData = Type.Missing;
        if (sourceData is MsExcel.Range rrange)
            comSourceData = new ExcelRange(rrange);
        else if (sourceData is MsExcel.ListObject lo)
            comSourceData = new ExcelListObject(lo);
        else if (sourceData is MsExcel.PivotTable dt)
            comSourceData = new ExcelPivotTable(dt);
        else if (sourceData is string sourceString)
            comSourceData = sourceString;
        return comSourceData;
    }

    public XlPivotTableVersionList Version => _pivotTable?.Version.EnumConvert(XlPivotTableVersionList.xlPivotTableVersionCurrent) ?? XlPivotTableVersionList.xlPivotTableVersionCurrent;

    #endregion

    #region 数据和字段
    public IExcelRange? DataBodyRange => _pivotTable != null ? new ExcelRange(_pivotTable.DataBodyRange) : null;
    public IExcelRange? TableRange1 => _pivotTable != null ? new ExcelRange(_pivotTable.TableRange1) : null;
    public IExcelRange? TableRange2 => _pivotTable != null ? new ExcelRange(_pivotTable.TableRange2) : null;
    public IExcelPivotFields? PageFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.PageFields) : null;
    public IExcelPivotFields? RowFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.RowFields) : null;
    public IExcelPivotFields? ColumnFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.ColumnFields) : null;
    public IExcelPivotFields? DataFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.DataFields) : null;
    public IExcelPivotFields? VisibleFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.VisibleFields) : null;
    public IExcelPivotFields? HiddenFields => _pivotTable != null ? new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.HiddenFields) : null;
    #endregion

    #region 格式和布局
    public IExcelTableStyle TableStyle
    {
        get => new ExcelTableStyle(_pivotTable.TableStyle2 as MsExcel.TableStyle);
        set
        {
            if (_pivotTable != null)
            {
                _pivotTable.TableStyle2 = ((ExcelTableStyle)value)._tableStyle;
            }
        }
    }

    public bool ShowTableStyleRowStripes
    {
        get => _pivotTable?.ShowTableStyleRowStripes ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.ShowTableStyleRowStripes = value;
        }
    }

    public bool ShowTableStyleColumnStripes
    {
        get => _pivotTable?.ShowTableStyleColumnStripes ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.ShowTableStyleColumnStripes = value;
        }
    }

    public bool ShowTableStyleLastColumn
    {
        get => _pivotTable?.ShowTableStyleLastColumn ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.ShowTableStyleLastColumn = value;
        }
    }

    public bool RowGrand
    {
        get => _pivotTable?.RowGrand ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.RowGrand = value;
        }
    }

    public bool ColumnGrand
    {
        get => _pivotTable?.ColumnGrand ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.ColumnGrand = value;
        }
    }

    public bool HasAutoFormat
    {
        get => _pivotTable?.HasAutoFormat ?? false;
        set
        {
            if (_pivotTable != null)
                _pivotTable.HasAutoFormat = value;
        }
    }
    #endregion

    #region 状态属性
    public bool IsProtected => !_pivotTable.EnableWizard || !_pivotTable.EnableDataValueEditing; // Simplified check
    #endregion

    #region 操作方法

    public IExcelPivotField? PivotFields(object Index)
    {
        if (_pivotTable == null) return null;
        var pf = _pivotTable.PivotFields(Index);
        if (pf is MsExcel.PivotField field)
            return new ExcelPivotField(field);
        return null;
    }

    public IExcelPivotFields? PivotFields()
    {
        if (_pivotTable == null) return null;
        var pf = _pivotTable.PivotFields();
        if (pf is MsExcel.PivotFields fields)
            return new ExcelPivotFields(fields);
        return null;
    }

    public IExcelCalculatedFields? CalculatedFields()
    {
        if (_pivotTable == null) return null;
        var calFields = _pivotTable.CalculatedFields();
        if (calFields is MsExcel.CalculatedFields fields)
            return new ExcelCalculatedFields(fields);
        return null;
    }

    public void Select(bool replace = true)
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.Select();
    }

    public void Copy()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.Copy();
    }

    public void Cut()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.Cut();
    }

    public void Delete()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.Clear();
    }
    #endregion

    #region 数据透视表操作
    public void Refresh()
    {
        if (_pivotTable == null) return;
        _pivotTable.RefreshTable();
    }

    public void Update()
    {
        Refresh();
    }

    public void Clear()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.ClearContents();
    }

    public void ClearFormats()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.ClearFormats();
    }

    public void ClearAll()
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.Clear();
    }

    public void ApplyAutoFormat(XlRangeAutoFormat format = XlRangeAutoFormat.xlRangeAutoFormatClassic1)
    {
        if (_pivotTable == null) return;
        _pivotTable.TableRange1.AutoFormat(format.EnumConvert(MsExcel.XlRangeAutoFormat.xlRangeAutoFormatClassic1));
    }
    #endregion

    #region 格式设置
    public void SetStyle(string styleName)
    {
        if (_pivotTable == null) return;
        _pivotTable.TableStyle2 = styleName;
    }
    #endregion

    #region 高级功能
    public void PrintOut(bool preview = false)
    {
        if (_pivotTable == null) return;
        if (preview)
        {
            _pivotTable.TableRange1.PrintPreview();
        }
        else
        {
            _pivotTable.TableRange1.PrintOutEx();
        }
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_pivotTable != null)
                Marshal.ReleaseComObject(_pivotTable);
            _pivotTable = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotTable()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
