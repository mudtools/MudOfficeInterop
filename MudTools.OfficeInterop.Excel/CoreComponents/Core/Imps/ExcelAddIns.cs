//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelAddIns : IExcelAddIns
{
    private MsExcel.AddIns _addIns;
    private bool _disposedValue;

    public object? Parent => _addIns.Parent;

    public int Count => _addIns.Count;

    public object Application => _addIns.Application;

    public IExcelAddIn this[object index] => new ExcelAddIn(_addIns[index]);

    internal ExcelAddIns(MsExcel.AddIns addIns)
    {
        _addIns = addIns ?? throw new ArgumentNullException(nameof(addIns));
        _disposedValue = false;
    }

    public IExcelAddIn Add(string filename, object copyFile)
    {
        if (string.IsNullOrEmpty(filename))
            throw new ArgumentException("文件名不能为空。", nameof(filename));

        try
        {
            var addIn = _addIns.Add(filename, copyFile);
            return addIn != null ? new ExcelAddIn(addIn) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加加载项: {filename}", ex);
        }
    }

    public IEnumerator<IExcelAddIn> GetEnumerator()
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

        if (disposing && _addIns != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_addIns) > 0) { }
            }
            catch { }
            _addIns = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}