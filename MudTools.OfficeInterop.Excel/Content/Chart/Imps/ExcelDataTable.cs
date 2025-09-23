//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel DataTable 对象的二次封装实现类
/// 实现 IExcelDataTable 接口
/// </summary>
internal class ExcelDataTable : IExcelDataTable
{
    private MsExcel.DataTable? _dataTable;
    private bool _disposedValue = false;

    internal ExcelDataTable(MsExcel.DataTable dataTable)
    {
        _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));
    }

    #region 基础属性  
    public object? Parent => _dataTable != null ? _dataTable.Parent : null;

    public IExcelApplication? Application => _dataTable != null ? new ExcelApplication(_dataTable.Application) : null;
    #endregion

    #region 格式设置
    public IExcelFont? Font
    {
        get
        {
            if (_dataTable == null)
                return null;
            return new ExcelFont(_dataTable.Font);
        }
    }

    public IExcelBorder? Border
    {
        get
        {
            if (_dataTable == null)
                return null;
            return new ExcelBorder(_dataTable.Border);
        }
    }

    public IExcelChartFormat? Format
    {
        get
        {
            if (_dataTable == null)
                return null;
            return new ExcelChartFormat(_dataTable.Format);
        }
    }

    public bool AutoScaleFont
    {
        get => _dataTable != null ? _dataTable.AutoScaleFont.ConvertToBool() : false;
        set
        {
            if (_dataTable == null)
                return;
            _dataTable.AutoScaleFont = value;
        }
    }

    public bool ShowLegendKey
    {
        get => _dataTable != null ? _dataTable.ShowLegendKey : false;
        set
        {
            if (_dataTable == null)
                return;
            _dataTable.ShowLegendKey = value;
        }
    }

    public bool HasBorderHorizontal
    {
        get => _dataTable != null ? _dataTable.HasBorderHorizontal : false;
        set
        {
            if (_dataTable == null)
                return;
            _dataTable.HasBorderHorizontal = value;
        }
    }

    public bool HasBorderVertical
    {
        get => _dataTable != null ? _dataTable.HasBorderVertical : false;
        set
        {
            if (_dataTable == null)
                return;
            _dataTable.HasBorderVertical = value;
        }
    }

    public bool HasBorderOutline
    {
        get => _dataTable != null ? _dataTable.HasBorderOutline : false;
        set
        {
            if (_dataTable == null)
                return;
            _dataTable.HasBorderOutline = value;
        }
    }
    #endregion

    #region 操作方法
    public void Select()
    {
        _dataTable?.Select();
    }

    public void Delete()
    {
        _dataTable?.Delete();
    }

    #endregion

    #region 格式设置方法   

    public void SetDataTableStyle(bool hasHorizontal, bool hasVertical, bool hasOutline)
    {
        HasBorderHorizontal = hasHorizontal;
        HasBorderVertical = hasVertical;
        HasBorderOutline = hasOutline;
    }
    #endregion    

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_dataTable != null)
                Marshal.ReleaseComObject(_dataTable);
            _dataTable = null;
        }
        _disposedValue = true;
    }

    ~ExcelDataTable()
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
