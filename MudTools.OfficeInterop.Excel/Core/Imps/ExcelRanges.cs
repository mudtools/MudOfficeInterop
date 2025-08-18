//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelRanges : IExcelRanges
{
    private MsExcel.Ranges _ranges;
    private bool _disposedValue;

    public int Count => _ranges.Count;

    public IExcelRange this[int index] => new ExcelRange(_ranges[index]);

    public IExcelRange this[string name] => new ExcelRange(_ranges.Item[name]);

    internal ExcelRanges(MsExcel.Ranges ranges)
    {
        _ranges = ranges ?? throw new ArgumentNullException(nameof(ranges));
        _disposedValue = false;
    }

    public IExcelRange GetItem(object index)
    {
        try
        {
            var range = _ranges[index];
            return range != null ? new ExcelRange(range) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public IExcelRange GetRange(
        string address,
        XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
        bool external = false)
    {
        if (string.IsNullOrEmpty(address))
            throw new ArgumentException("范围地址不能为空。", nameof(address));

        try
        {
            var sheet = _ranges.Parent as MsExcel.Worksheet;
            var range = sheet.Range[address];
            return range != null ? new ExcelRange(range) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法获取范围: {address}", ex);
        }
    }


    public IExcelWorksheet Parent => new ExcelWorksheet(_ranges.Parent as MsExcel.Worksheet); 

    public IEnumerator<IExcelRange> GetEnumerator()
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

        if (disposing && _ranges != null)
        {
            try
            {
                Marshal.ReleaseComObject(_ranges);
            }
            catch { }
            _ranges = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}