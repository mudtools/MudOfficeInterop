//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelSortFields : IExcelSortFields
{
    private MsExcel.SortFields _sortFields;
    private bool _disposedValue;

    public int Count => _sortFields.Count;

    public IExcelApplication Application
    {
        get
        {
            var application = _sortFields?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public IExcelSortField this[int index] => new ExcelSortField(_sortFields[index]);

    internal ExcelSortFields(MsExcel.SortFields sortFields)
    {
        _sortFields = sortFields ?? throw new ArgumentNullException(nameof(sortFields));
        _disposedValue = false;
    }

    public IExcelSortField? Add(IExcelRange key, XlSortOn sortOn = XlSortOn.xlSortOnValues,
                              XlSortOrder order = XlSortOrder.xlAscending,
                              object? customOrder = null, XlSortDataOption dataOption = XlSortDataOption.xlSortNormal)
    {
        if (key == null) throw new ArgumentNullException(nameof(key));

        try
        {
            MsExcel.Range? comRange = null;
            if (key is ExcelRange excelRange)
            {
                comRange = excelRange.InternalRange;
            }

            object customOrderObj = customOrder ?? Type.Missing;

            var sortField = _sortFields.Add(comRange,
                                          (MsExcel.XlSortOn)sortOn,
                                          (MsExcel.XlSortOrder)order,
                                          customOrderObj,
                                          (MsExcel.XlSortDataOption)dataOption);

            return sortField != null ? new ExcelSortField(sortField) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加排序字段。", ex);
        }
    }

    public void Clear()
    {
        try
        {
            _sortFields.Clear();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法清除排序字段。", ex);
        }
    }

    public IExcelSort Parent => new ExcelSort(_sortFields.Parent as MsExcel.Sort);

    public void RemoveAt(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        try
        {
            this[index].Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法移除索引为 {index} 的排序字段。", ex);
        }
    }

    public IEnumerator<IExcelSortField> GetEnumerator()
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

        if (disposing && _sortFields != null)
        {
            try
            {
                Marshal.ReleaseComObject(_sortFields);
            }
            catch { }
            _sortFields = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}