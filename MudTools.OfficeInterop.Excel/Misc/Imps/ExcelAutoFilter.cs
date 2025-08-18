//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelAutoFilter : IExcelAutoFilter
{
    private MsExcel.AutoFilter _autoFilter;
    private bool _disposedValue;

    public bool FilterMode => _autoFilter.FilterMode;

    public IExcelApplication Application
    {
        get
        {
            var application = _autoFilter?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public IExcelRange Range => new ExcelRange(_autoFilter.Range);

    public object Parent => _autoFilter.Parent;

    public IExcelFilters Filters => new ExcelFilters(_autoFilter.Filters);

    public IExcelSort Sort => new ExcelSort(_autoFilter.Sort);

    internal ExcelAutoFilter(MsExcel.AutoFilter autoFilter)
    {
        _autoFilter = autoFilter ?? throw new ArgumentNullException(nameof(autoFilter));
        _disposedValue = false;
    }

    public void ApplyFilter()
    {
        try
        {
            _autoFilter.ApplyFilter();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除筛选条件。", ex);
        }
    }


    public void ShowAllData()
    {
        try
        {
            _autoFilter.ShowAllData();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除筛选条件。", ex);
        }
    }


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _autoFilter != null)
        {
            try
            {
                Marshal.ReleaseComObject(_autoFilter);
            }
            catch { }
            _autoFilter = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}