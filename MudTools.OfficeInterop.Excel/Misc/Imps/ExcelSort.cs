//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelSort : IExcelSort
{
    private MsExcel.Sort _sort;
    private bool _disposedValue;

    internal ExcelSort(MsExcel.Sort sort)
    {
        _sort = sort ?? throw new ArgumentNullException(nameof(sort));
        _disposedValue = false;
    }

    public IExcelRange Range
    {
        get => new ExcelRange(_sort.Rng);
    }

    public IExcelApplication Application
    {
        get
        {
            var application = _sort?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public XlYesNoGuess Header
    {
        get => (XlYesNoGuess)(int)_sort.Header;
        set => _sort.Header = (MsExcel.XlYesNoGuess)value;
    }


    public XlSortMethod SortMethod
    {
        get => (XlSortMethod)(int)_sort.SortMethod;
        set => _sort.SortMethod = (MsExcel.XlSortMethod)value;
    }

    public IExcelSortFields SortFields => new ExcelSortFields(_sort.SortFields);

    public object Parent => _sort.Parent;


    public bool MatchCase
    {
        get => _sort.MatchCase;
        set => _sort.MatchCase = value;
    }

    public XlSortOrientation Orientation
    {
        get => (XlSortOrientation)(int)_sort.Orientation;
        set => _sort.Orientation = (MsExcel.XlSortOrientation)value;
    }

    public void SetRange(IExcelRange range)
    {
        if (range == null)
            return;
        MsExcel.Range r = ((ExcelRange)range).InternalRange;
        _sort.SetRange(r);
    }
    public void Apply()
    {
        try
        {
            _sort.Apply();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法应用排序。", ex);
        }
    }


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _sort != null)
        {
            try
            {
                Marshal.ReleaseComObject(_sort);
            }
            catch { }
            _sort = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}