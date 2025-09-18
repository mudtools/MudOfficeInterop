//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelListObjects : IExcelListObjects
{
    private MsExcel.ListObjects _listObjects;
    private bool _disposedValue;

    public object Parent => _listObjects.Parent;

    public int Count => _listObjects.Count;

    public object Application => _listObjects.Application;


    public IExcelListObject this[object index] => new ExcelListObject(_listObjects[index]);

    internal ExcelListObjects(MsExcel.ListObjects listObjects)
    {
        _listObjects = listObjects ?? throw new ArgumentNullException(nameof(listObjects));
        _disposedValue = false;
    }

    public IExcelListObject Add(XlListObjectSourceType sourceType, object source, object link, XlYesNoGuess xlListObjectHasHeaders = XlYesNoGuess.xlGuess, object? destination = null)
    {
        try
        {
            var listObject = _listObjects.Add((MsExcel.XlListObjectSourceType)(int)sourceType, source ?? Type.Missing, link ?? Type.Missing, (MsExcel.XlYesNoGuess)(int)xlListObjectHasHeaders, destination ?? Type.Missing);
            return listObject != null ? new ExcelListObject(listObject) : null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加列表对象。", ex);
        }
    }

    public IEnumerator<IExcelListObject> GetEnumerator()
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

        if (disposing && _listObjects != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_listObjects) > 0) { }
            }
            catch { }
            _listObjects = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}