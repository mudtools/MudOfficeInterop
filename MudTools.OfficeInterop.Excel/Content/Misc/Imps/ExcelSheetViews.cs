//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelSheetViews : IExcelSheetViews
{
    private MsExcel.SheetViews _sheetViews;
    private bool _disposedValue;

    public object Parent => _sheetViews.Parent;

    public int Count => _sheetViews.Count;

    public object Application => _sheetViews.Application;



    public IExcelUserAccess this[object index]
    {
        get
        {
            try
            {
                MsExcel.UserAccess? sheetView = _sheetViews[index] as MsExcel.UserAccess;
                return sheetView != null ? new ExcelUserAccess(sheetView) : null;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法获取工作表视图项目。", ex);
            }
        }
    }

    internal ExcelSheetViews(MsExcel.SheetViews sheetViews)
    {
        _sheetViews = sheetViews ?? throw new ArgumentNullException(nameof(sheetViews));
        _disposedValue = false;
    }


    public IEnumerator<IExcelUserAccess> GetEnumerator()
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

        if (disposing && _sheetViews != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_sheetViews) > 0) { }
            }
            catch { }
            _sheetViews = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}