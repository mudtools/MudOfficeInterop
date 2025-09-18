//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelPage : IExcelPage
{
    private MsExcel.Page _page;
    private bool _disposedValue;

    internal ExcelPage(MsExcel.Page headerFooter)
    {
        _page = headerFooter ?? throw new ArgumentNullException(nameof(headerFooter));

        _disposedValue = false;
    }

    public IExcelHeaderFooter? LeftHeader => _page != null ? new ExcelHeaderFooter(_page.LeftHeader) : null;

    public IExcelHeaderFooter? CenterHeader => _page != null ? new ExcelHeaderFooter(_page.CenterHeader) : null;

    public IExcelHeaderFooter? RightHeader => _page != null ? new ExcelHeaderFooter(_page.RightHeader) : null;

    public IExcelHeaderFooter? LeftFooter => _page != null ? new ExcelHeaderFooter(_page.LeftFooter) : null;

    public IExcelHeaderFooter? CenterFooter => _page != null ? new ExcelHeaderFooter(_page.CenterFooter) : null;

    public IExcelHeaderFooter? RightFooter => _page != null ? new ExcelHeaderFooter(_page.RightFooter) : null;


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _page != null)
        {
            try
            {
                Marshal.ReleaseComObject(_page);
            }
            catch { }
            _page = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
