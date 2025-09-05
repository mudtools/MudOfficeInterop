//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    private MsExcel.PivotTable? _pivotTable;
    private bool _disposedValue = false;

    internal ExcelPivotTable(MsExcel.PivotTable pivotTable)
    {
        _pivotTable = pivotTable ?? throw new ArgumentNullException(nameof(pivotTable));
    }

    #region 基础属性
    public string Name
    {
        get => _pivotTable.Name;
        set => _pivotTable.Name = value;
    }

    public object Parent => _pivotTable.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotTable.Application);

    public IExcelPivotCache PivotCache() => new ExcelPivotCache(_pivotTable.PivotCache());

    public object SourceData => _pivotTable.SourceData;

    public int Version => (int)_pivotTable.Version;

    #endregion

    #region 数据和字段
    public IExcelRange DataBodyRange => new ExcelRange(_pivotTable.DataBodyRange);
    public IExcelRange TableRange1 => new ExcelRange(_pivotTable.TableRange1);
    public IExcelRange TableRange2 => new ExcelRange(_pivotTable.TableRange2);
    public IExcelPivotFields PivotFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.PivotFields());
    public IExcelPivotFields PageFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.PageFields);
    public IExcelPivotFields RowFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.RowFields);
    public IExcelPivotFields ColumnFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.ColumnFields);
    public IExcelPivotFields DataFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.DataFields);
    public IExcelPivotFields VisibleFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.VisibleFields);
    public IExcelPivotFields HiddenFields => new ExcelPivotFields((MsExcel.PivotFields)_pivotTable.HiddenFields);
    #endregion

    #region 格式和布局
    public IExcelTableStyle TableStyle
    {
        get => new ExcelTableStyle(_pivotTable.TableStyle2 as MsExcel.TableStyle);
        set
        {
            throw new NotImplementedException();
        }
    }

    public bool ShowRowStripes
    {
        get => _pivotTable.ShowTableStyleRowStripes;
        set => _pivotTable.ShowTableStyleRowStripes = value;
    }

    public bool ShowColumnStripes
    {
        get => _pivotTable.ShowTableStyleColumnStripes;
        set => _pivotTable.ShowTableStyleColumnStripes = value;
    }

    public bool ShowLastColumn
    {
        get => _pivotTable.ShowTableStyleLastColumn;
        set => _pivotTable.ShowTableStyleLastColumn = value;
    }
    #endregion

    #region 状态属性
    public bool IsProtected => !_pivotTable.EnableWizard || !_pivotTable.EnableDataValueEditing; // Simplified check
    #endregion

    #region 操作方法
    public void Select(bool replace = true)
    {
        _pivotTable.TableRange1.Select();
    }

    public void Copy()
    {
        _pivotTable.TableRange1.Copy();
    }

    public void Cut()
    {
        _pivotTable.TableRange1.Cut();
    }

    public void Delete()
    {
        _pivotTable.TableRange1.Clear();
    }
    #endregion

    #region 数据透视表操作
    public void Refresh()
    {
        _pivotTable.RefreshTable();
    }

    public void Update()
    {
        Refresh();
    }

    public void Clear()
    {
        _pivotTable.TableRange1.ClearContents();
    }

    public void ClearFormats()
    {
        _pivotTable.TableRange1.ClearFormats();
    }

    public void ClearAll()
    {
        _pivotTable.TableRange1.Clear();
    }

    public void ApplyAutoFormat(int format = 1)
    {
        _pivotTable.TableRange1.AutoFormat((MsExcel.XlRangeAutoFormat)format);
    }
    #endregion

    #region 格式设置
    public void SetStyle(string styleName)
    {
        _pivotTable.TableStyle2 = styleName;
    }
    #endregion

    #region 高级功能
    public void PrintOut(bool preview = false)
    {
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
            try
            {
                // 释放底层COM对象
                if (_pivotTable != null)
                    Marshal.ReleaseComObject(_pivotTable);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
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
