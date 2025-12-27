//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Areas 集合对象的二次封装实现类
/// 实现 IExcelAreas 接口
/// </summary>
internal class ExcelAreas : IExcelAreas
{
    private MsExcel.Areas _areas;
    private bool _disposedValue = false;

    internal ExcelAreas(MsExcel.Areas areas)
    {
        _areas = areas ?? throw new ArgumentNullException(nameof(areas));
    }

    #region 基础属性
    public int Count => _areas.Count;

    public IExcelRange this[int index] => new ExcelRange((MsExcel.Range)_areas[index]);

    public object? Parent => _areas.Parent;

    public IExcelApplication? Application => new ExcelApplication(_areas.Application);
    #endregion

    #region 查找和筛选
    public IExcelRange[] FindByAddress(string address, bool matchCase = false)
    {
        var results = new List<IExcelRange>();
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            if (string.Compare(area.Address, address, !matchCase) == 0)
            {
                results.Add(area);
            }
        }
        return results.ToArray();
    }

    public IExcelRange[] FindBySize(int rowCount, int columnCount, int tolerance = 0)
    {
        var results = new List<IExcelRange>();
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            if (Math.Abs(area.Rows.Count - rowCount) <= tolerance &&
                Math.Abs(area.Columns.Count - columnCount) <= tolerance)
            {
                results.Add(area);
            }
        }
        return results.ToArray();
    }

    public IExcelRange GetLargestArea()
    {
        IExcelRange largest = null;
        double maxArea = 0;
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            double currentArea = area.Rows.Count * area.Columns.Count;
            if (currentArea > maxArea)
            {
                maxArea = currentArea;
                largest = area;
            }
        }
        return largest;
    }

    public IExcelRange GetSmallestArea()
    {
        IExcelRange smallest = null;
        double minArea = double.MaxValue;
        if (Count > 0) minArea = this[1].Rows.Count * this[1].Columns.Count + 1;
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            double currentArea = area.Rows.Count * area.Columns.Count;
            if (currentArea < minArea)
            {
                minArea = currentArea;
                smallest = area;
            }
        }
        return smallest;
    }

    public IExcelRange[] GetVisibleAreas()
    {
        var results = new List<IExcelRange>();
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            if (area.EntireRow.Hidden == false && area.EntireColumn.Hidden == false)
            {
                results.Add(area);
            }
        }
        return results.ToArray();
    }

    public IExcelRange[] GetHiddenAreas()
    {
        var results = new List<IExcelRange>();
        for (int i = 1; i <= Count; i++)
        {
            var area = this[i];
            if (area.EntireRow.Hidden == true || area.EntireColumn.Hidden == true)
            {
                results.Add(area);
            }
        }
        return results.ToArray();
    }
    #endregion

    #region 操作方法   

    public void Delete(int index)
    {
        try
        {
            MsExcel.Range areaRange = _areas[index] as MsExcel.Range;
            areaRange.Delete();
        }
        catch { /* Handle error */ }
    }

    public void Delete(IExcelRange area)
    {
        if (area is ExcelRange excelRange)
        {
            try
            {
                excelRange.InternalRange?.Delete();
            }
            catch { /* Handle error */ }
        }
    }

    public void DeleteRange(int[] indices)
    {
        // Sort indices descending to avoid index shifting issues if deleting cells
        var sortedIndices = new List<int>(indices);
        sortedIndices.Sort((a, b) => b.CompareTo(a)); // Descending sort
        foreach (int index in sortedIndices)
        {
            Delete(index);
        }
    }
    #endregion


    #region IEnumerable<IExcelRange> Support
    public IEnumerator<IExcelRange> GetEnumerator()
    {
        for (int i = 1; i <= _areas.Count; i++)
        {
            yield return new ExcelRange(_areas[i]);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {

                // 释放底层COM对象
                if (_areas != null)
                    Marshal.ReleaseComObject(_areas);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _areas = null;
        }

        _disposedValue = true;
    }

    ~ExcelAreas()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}