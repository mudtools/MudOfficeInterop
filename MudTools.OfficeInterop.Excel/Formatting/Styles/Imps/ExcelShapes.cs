

namespace MudTools.OfficeInterop.Excel.Imps;

internal partial class ExcelShapes : IExcelShapes
{
    public IExcelShapeRange? Range(object index)
    {
        if (_shapes == null)
            throw new ObjectDisposedException(nameof(_shapes));

        try
        {
            var comObj = _shapes?.Range[index];
            if (comObj == null)
                return null;
            return new ExcelShapeRange(comObj);
        }
        catch (COMException cx)
        {
            throw new InvalidOperationException("执行AddTextEffect操作失败。", cx);
        }
        catch (Exception ex)
        {
            throw new ExcelOperationException("执行AddTextEffect操作失败", ex);
        }
    }

}