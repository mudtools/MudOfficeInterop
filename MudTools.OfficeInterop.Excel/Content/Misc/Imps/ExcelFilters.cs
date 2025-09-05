//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelFilters : IExcelFilters
{
    private MsExcel.Filters _autoFilters;
    private bool _disposedValue;

    public int Count => _autoFilters.Count;

    public IExcelApplication Application
    {
        get
        {
            var application = _autoFilters?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public IExcelFilter this[int index] => new ExcelFilter(_autoFilters.Item[index]);

    internal ExcelFilters(MsExcel.Filters autoFilters)
    {
        _autoFilters = autoFilters ?? throw new ArgumentNullException(nameof(autoFilters));
        _disposedValue = false;
    }


    public bool ApplyFilters(IExcelRange range)
    {
        if (range == null) throw new ArgumentNullException(nameof(range));

        try
        {
            var comRange = ((ExcelRange)range).InternalRange;
            comRange.AutoFilter();
            return true;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法应用自动筛选。", ex);
        }
    }


    public object Parent => _autoFilters.Parent;


    public IEnumerator<IExcelFilter> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _autoFilters != null)
        {
            try
            {
                Marshal.ReleaseComObject(_autoFilters);
            }
            catch { }
            _autoFilters = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}