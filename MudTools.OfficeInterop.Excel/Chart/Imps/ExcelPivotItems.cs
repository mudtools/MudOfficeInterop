//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PivotItems 集合对象的二次封装实现类
/// 实现 IExcelPivotItems 接口
/// </summary>
internal class ExcelPivotItems : IExcelPivotItems
{
    private MsExcel.PivotItems _pivotItems;
    private bool _disposedValue = false;

    internal ExcelPivotItems(MsExcel.PivotItems pivotItems)
    {
        _pivotItems = pivotItems ?? throw new ArgumentNullException(nameof(pivotItems));
    }

    #region 基础属性
    public int Count => _pivotItems.Count;

    public IExcelPivotItem this[int index]
    {
        get
        {
            if (_pivotItems == null || index < 1 || index > Count)
                return null;

            try
            {
                var pivotObject = _pivotItems.Item(index) as MsExcel.PivotItem;
                return pivotObject != null ? new ExcelPivotItem(pivotObject) : null;
            }
            catch
            {
                return null;
            }
        }
    }
    public IExcelPivotItem this[string name]
    {
        get
        {
            if (_pivotItems == null || string.IsNullOrEmpty(name))
                return null;

            try
            {
                var result = this.FindByName(name);
                if (result != null && result.Length > 0)
                    return result[0];
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    public object Parent => _pivotItems.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotItems.Application);
    #endregion

    #region 查找和筛选
    public IExcelPivotItem[] FindByName(string name, bool matchCase = false)
    {
        List<IExcelPivotItem> results = [];
        for (int i = 1; i <= Count; i++)
        {
            IExcelPivotItem item = this[i];
            if (string.Compare(item.Name, name, !matchCase) == 0)
            {
                results.Add(item);
            }
        }
        return results.ToArray();
    }

    public IExcelPivotItem[] FindByVisibility(bool isVisible)
    {
        var results = new List<IExcelPivotItem>();
        for (int i = 1; i <= Count; i++)
        {
            var item = this[i];
            if (item.Visible == isVisible)
            {
                results.Add(item);
            }
        }
        return results.ToArray();
    }


    public IExcelPivotItem[] GetVisibleItems()
    {
        return FindByVisibility(true);
    }

    public IExcelPivotItem[] GetHiddenItems()
    {
        return FindByVisibility(false);
    }
    #endregion

    #region 操作方法
    public void HideAll()
    {
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                this[i].Visible = false;
            }
            catch
            {
                // Handle error, e.g., if item cannot be hidden individually
            }
        }
    }

    public void ShowAll()
    {
        for (int i = 1; i <= Count; i++)
        {
            this[i].Visible = true;
        }
    }

    public void Hide(int index)
    {
        try
        {
            this[index].Visible = false;
        }
        catch
        {
            // Handle error if index is invalid
        }
    }

    public void Hide(string name)
    {
        try
        {
            this[name].Visible = false;
        }
        catch
        {
            // Handle error if name is invalid
        }
    }

    public void Hide(IExcelPivotItem item)
    {
        if (item is ExcelPivotItem excelItem)
        {
            try
            {
                excelItem._pivotItem.Visible = false;
            }
            catch
            {
                // Handle error
            }
        }
    }

    public void HideRange(int[] indices)
    {
        foreach (int index in indices)
        {
            Hide(index);
        }
    }
    #endregion

    #region IEnumerable<IExcelPivotItem> Support
    public IEnumerator<IExcelPivotItem> GetEnumerator()
    {
        for (int i = 1; i <= _pivotItems.Count; i++)
        {
            yield return new ExcelPivotItem((MsExcel.PivotItem)_pivotItems.Item(i));
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
                if (_pivotItems != null)
                    Marshal.ReleaseComObject(_pivotItems);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pivotItems = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotItems()
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
