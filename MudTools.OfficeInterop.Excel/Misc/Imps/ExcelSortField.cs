//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelSortField : IExcelSortField
{
    private MsExcel.SortField _sortField;
    private bool _disposedValue;

    public IExcelApplication Application
    {
        get
        {
            var application = _sortField?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public IExcelRange Key
    {
        get => new ExcelRange(_sortField.Key);
    }

    public XlSortOn SortOn
    {
        get => (XlSortOn)(int)_sortField.SortOn;
        set => _sortField.SortOn = (MsExcel.XlSortOn)value;
    }

    public XlSortOrder Order
    {
        get => (XlSortOrder)(int)_sortField.Order;
        set => _sortField.Order = (MsExcel.XlSortOrder)value;
    }

    public object CustomOrder
    {
        get => _sortField.CustomOrder;
        set => _sortField.CustomOrder = value ?? Type.Missing;
    }

    public XlSortDataOption DataOption
    {
        get => (XlSortDataOption)(int)_sortField.DataOption;
        set => _sortField.DataOption = (MsExcel.XlSortDataOption)value;
    }

    public object SortOnValue
    {
        get => _sortField.SortOnValue;
    }

    public IExcelSortFields Parent => new ExcelSortFields(_sortField.Parent as MsExcel.SortFields);

    public int Priority
    {
        get => _sortField.Priority;
        set => _sortField.Priority = value;
    }


    internal ExcelSortField(MsExcel.SortField sortField)
    {
        _sortField = sortField ?? throw new ArgumentNullException(nameof(sortField));
        _disposedValue = false;
    }

    public void Delete()
    {
        try
        {
            _sortField.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除排序字段。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _sortField != null)
        {
            try
            {
                Marshal.ReleaseComObject(_sortField);
            }
            catch { }
            _sortField = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}