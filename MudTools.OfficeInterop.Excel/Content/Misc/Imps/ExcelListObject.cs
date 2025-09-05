//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelListObject : IExcelListObject
{
    private MsExcel.ListObject _listObject;
    private bool _disposedValue;

    public string Name
    {
        get => _listObject.Name;
        set => _listObject.Name = value;
    }
    public IExcelRange Range => new ExcelRange(_listObject.Range);
    public IExcelRange DataRange => new ExcelRange(_listObject.DataBodyRange);

    public IExcelRange HeaderRowRange => new ExcelRange(_listObject.HeaderRowRange);

    public IExcelRange TotalsRowRange => new ExcelRange(_listObject.TotalsRowRange);

    public bool ShowHeaders
    {
        get => _listObject.ShowHeaders;
        set => _listObject.ShowHeaders = value;
    }

    public bool ShowTotals
    {
        get => _listObject.ShowTotals;
        set => _listObject.ShowTotals = value;
    }

    public string WorksheetName => _listObject.Range.Worksheet.Name;

    internal ExcelListObject(MsExcel.ListObject listObject)
    {
        _listObject = listObject ?? throw new ArgumentNullException(nameof(listObject));
        _disposedValue = false;
    }

    public void Refresh()
    {
        try
        {
            _listObject.Refresh();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法刷新 ListObject。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _listObject.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除 ListObject。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listObject != null)
        {
            Marshal.ReleaseComObject(_listObject);
            _listObject = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}