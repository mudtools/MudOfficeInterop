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
    public IExcelRange? Range => _listObject != null ? new ExcelRange(_listObject.Range) : null;
    public IExcelRange? DataRange => _listObject != null ? new ExcelRange(_listObject.DataBodyRange) : null;

    public IExcelRange? HeaderRowRange => _listObject != null ? new ExcelRange(_listObject.HeaderRowRange) : null;

    public IExcelRange? TotalsRowRange => _listObject != null ? new ExcelRange(_listObject.TotalsRowRange) : null;

    public IExcelRange? InsertRowRange => _listObject != null ? new ExcelRange(_listObject.InsertRowRange) : null;

    public IExcelListColumns? ListColumns => _listObject != null ? new ExcelListColumns(_listObject.ListColumns) : null;

    public IExcelListRows? ListRows => _listObject != null ? new ExcelListRows(_listObject.ListRows) : null;

    public IExcelAutoFilter? AutoFilter => _listObject != null ? new ExcelAutoFilter(_listObject.AutoFilter) : null;

    public IExcelSort? Sort => _listObject != null ? new ExcelSort(_listObject.Sort) : null;

    public IExcelQueryTable? QueryTable => _listObject != null ? new ExcelQueryTable(_listObject.QueryTable) : null;

    public bool DisplayRightToLeft
    {
        get => _listObject.DisplayRightToLeft;
    }

    public bool ShowHeaders
    {
        get => _listObject.ShowHeaders;
        set => _listObject.ShowHeaders = value;
    }

    public bool ShowAutoFilter
    {
        get => _listObject.ShowAutoFilter;
        set => _listObject.ShowAutoFilter = value;
    }

    public bool ShowTableStyleFirstColumn
    {
        get => _listObject.ShowTableStyleFirstColumn;
        set => _listObject.ShowTableStyleFirstColumn = value;
    }

    public bool ShowTableStyleLastColumn
    {
        get => _listObject.ShowTableStyleLastColumn;
        set => _listObject.ShowTableStyleLastColumn = value;
    }

    public bool ShowTableStyleRowStripes
    {
        get => _listObject.ShowTableStyleRowStripes;
        set => _listObject.ShowTableStyleRowStripes = value;
    }

    public bool ShowTableStyleColumnStripes
    {
        get => _listObject.ShowTableStyleColumnStripes;
        set => _listObject.ShowTableStyleColumnStripes = value;
    }

    public string SharePointURL
    {
        get => _listObject.SharePointURL;
    }

    public bool ShowTotals
    {
        get => _listObject.ShowTotals;
        set => _listObject.ShowTotals = value;
    }

    public string DisplayName
    {
        get => _listObject.DisplayName;
        set => _listObject.DisplayName = value;
    }

    public string Comment
    {
        get => _listObject.Comment;
        set => _listObject.Comment = value;
    }

    public string AlternativeText
    {
        get => _listObject.AlternativeText;
        set => _listObject.AlternativeText = value;
    }

    public string Summary
    {
        get => _listObject.Summary;
        set => _listObject.Summary = value;
    }

    public bool ShowAutoFilterDropDown
    {
        get => _listObject.ShowAutoFilterDropDown;
        set => _listObject.ShowAutoFilterDropDown = value;
    }

    public XlListObjectSourceType SourceType
    {
        get => _listObject.SourceType.EnumConvert(XlListObjectSourceType.xlSrcRange);
    }

    public string WorksheetName => _listObject.Range.Worksheet.Name;

    internal ExcelListObject(MsExcel.ListObject listObject)
    {
        _listObject = listObject ?? throw new ArgumentNullException(nameof(listObject));
        _disposedValue = false;
    }

    public void ExportToVisio()
    {
        try
        {
            _listObject.ExportToVisio();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法将 ListObject 导出到 Visio。", ex);
        }
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