//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel PivotTables 集合对象的二次封装实现类
/// 实现 IExcelPivotTables 接口
/// </summary>
internal class ExcelPivotTables : IExcelPivotTables
{
    private MsExcel.PivotTables _pivotTables;
    private bool _disposedValue = false;

    internal ExcelPivotTables(MsExcel.PivotTables pivotTables)
    {
        _pivotTables = pivotTables ?? throw new ArgumentNullException(nameof(pivotTables));
    }

    #region 基础属性
    public int Count => _pivotTables.Count;

    public IExcelPivotTable this[int index] => new ExcelPivotTable((MsExcel.PivotTable)_pivotTables.Item(index));

    public IExcelPivotTable this[string name]
    {
        get
        {
            var obj = FindByName(name);
            return obj != null && obj.Length > 0 ? obj[0] : null;
        }
    }

    public object Parent => _pivotTables.Parent;

    public IExcelApplication Application => new ExcelApplication();
    #endregion

    #region 查找和筛选
    public IExcelPivotTable[] FindByName(string name, bool matchCase = false)
    {
        var results = new List<IExcelPivotTable>();
        for (int i = 1; i <= Count; i++)
        {
            var pt = this[i];
            if (string.Compare(pt.Name, name, !matchCase) == 0)
            {
                results.Add(pt);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotTable[] FindBySourceData(IExcelRange sourceData)
    {
        var results = new List<IExcelPivotTable>();
        for (int i = 1; i <= Count; i++)
        {
            var pt = this[i];
            if (pt.SourceData.Equals(sourceData))
            {
                results.Add(pt);
            }
        }
        return results.ToArray();
    }


    public IExcelPivotTable[] GetProtectedPivotTables()
    {
        var results = new List<IExcelPivotTable>();
        for (int i = 1; i <= Count; i++)
        {
            var pt = this[i];
            if (pt.IsProtected)
            {
                results.Add(pt);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotTable[] GetUnprotectedPivotTables()
    {
        var results = new List<IExcelPivotTable>();
        for (int i = 1; i <= Count; i++)
        {
            var pt = this[i];
            if (!pt.IsProtected)
            {
                results.Add(pt);
            }
        }
        return results.ToArray();
    }
    #endregion



    #region 高级功能
    public void PrintOutAll(bool preview = false)
    {
        for (int i = 1; i <= Count; i++)
        {
            this[i].PrintOut(preview);
        }
    }

    public void UpdateAll()
    {
        for (int i = 1; i <= Count; i++)
        {
            this[i].Update();
        }
    }
    #endregion

    #region IEnumerable<IExcelPivotTable> Support
    public IEnumerator<IExcelPivotTable> GetEnumerator()
    {
        for (int i = 1; i <= _pivotTables.Count; i++)
        {
            yield return new ExcelPivotTable((MsExcel.PivotTable)_pivotTables.Item(i));
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
        if (!_disposedValue)
        {
            if (_pivotTables != null)
            {
                Marshal.ReleaseComObject(_pivotTables);
                _pivotTables = null;
            }
            _disposedValue = true;
        }
    }

    ~ExcelPivotTables()
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
