//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelPane : IExcelPane
{
    private MsExcel.Pane _pane;
    private bool _disposedValue;

    public int Index => _pane.Index;

    public IExcelRange VisibleRange => new ExcelRange(_pane.VisibleRange);

    public int ScrollColumn
    {
        get => _pane.ScrollColumn;
        set => _pane.ScrollColumn = value;
    }

    public int ScrollRow
    {
        get => _pane.ScrollRow;
        set => _pane.ScrollRow = value;
    }

    internal ExcelPane(MsExcel.Pane pane)
    {
        _pane = pane ?? throw new ArgumentNullException(nameof(pane));
        _disposedValue = false;
    }

    public void Activate()
    {
        try
        {
            _pane.Activate();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法激活窗格。", ex);
        }
    }

    public void ScrollIntoView(IExcelRange range)
    {
        if (range == null) throw new ArgumentNullException(nameof(range));

        try
        {
            var comRange = range.RangeRect;
            _pane.ScrollIntoView(comRange.Left, comRange.Top, comRange.Width, comRange.Height);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法滚动到指定范围。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pane != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_pane) > 0) { }
            }
            catch { }
            _pane = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
