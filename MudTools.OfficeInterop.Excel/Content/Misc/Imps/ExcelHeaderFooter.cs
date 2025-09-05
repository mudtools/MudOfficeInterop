//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelHeaderFooter : IExcelHeaderFooter
{
    private MsExcel.HeaderFooter _headerFooter;
    private bool _disposedValue;

    internal ExcelHeaderFooter(MsExcel.HeaderFooter headerFooter)
    {
        _headerFooter = headerFooter ?? throw new ArgumentNullException(nameof(headerFooter));

        _disposedValue = false;
    }

    public string? Text
    {
        get => _headerFooter?.Text;
        set => _headerFooter.Text = value;
    }

    public IExcelGraphic? Picture => _headerFooter != null ? new ExcelGraphic(_headerFooter.Picture) : null;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _headerFooter != null)
        {
            try
            {
                Marshal.ReleaseComObject(_headerFooter);
            }
            catch { }
            _headerFooter = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
