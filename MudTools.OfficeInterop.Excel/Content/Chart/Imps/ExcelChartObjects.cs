//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ChartObjects 集合对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.ChartObjects 对象的安全访问和资源管理
/// </summary>
internal class ExcelChartObjects : IExcelChartObjects
{
    /// <summary>
    /// 底层的 COM ChartObjects 集合对象
    /// </summary>
    private MsExcel.ChartObjects _chartObjects;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelChartObjects 实例
    /// </summary>
    /// <param name="chartObjects">底层的 COM ChartObjects 集合对象</param>
    internal ExcelChartObjects(MsExcel.ChartObjects chartObjects)
    {
        _chartObjects = chartObjects;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放所有子图表对象
                for (int i = 1; i <= Count; i++)
                {
                    var chartObject = this[i] as ExcelChartObject;
                    chartObject?.Dispose();
                }

                // 释放底层COM对象
                if (_chartObjects != null)
                    Marshal.ReleaseComObject(_chartObjects);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _chartObjects = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取图表对象集合中的图表数量
    /// </summary>
    public int Count => _chartObjects?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的图表对象
    /// </summary>
    /// <param name="index">图表对象索引（从1开始）</param>
    /// <returns>图表对象</returns>
    public IExcelChartObject this[int index]
    {
        get
        {
            if (_chartObjects == null || index < 1 || index > Count)
                return null;

            try
            {
                var chartObject = _chartObjects.Item(index) as MsExcel.ChartObject;
                return chartObject != null ? new ExcelChartObject(chartObject) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取指定名称的图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>图表对象</returns>
    public IExcelChartObject this[string name]
    {
        get
        {
            if (_chartObjects == null || string.IsNullOrEmpty(name))
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

    /// <summary>
    /// 获取图表对象集合所在的父对象
    /// </summary>
    public object Parent => _chartObjects?.Parent;

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向工作表添加新的图表对象
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图表对象</returns>
    public IExcelChartObject Add(double left, double top, double width, double height)
    {
        if (_chartObjects == null)
            return null;

        try
        {
            var chartObject = _chartObjects.Add(left, top, width, height) as MsExcel.ChartObject;
            return chartObject != null ? new ExcelChartObject(chartObject) : null;
        }
        catch
        {
            return null;
        }
    }
    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>匹配的图表对象数组</returns>
    public IExcelChartObject[] FindByName(string name)
    {
        if (_chartObjects == null || string.IsNullOrEmpty(name) || Count == 0)
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null && chartObject.Name?.Contains(name) == true)
                {
                    result.Add(chartObject);
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据位置查找图表对象
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    public IExcelChartObject[] FindByPosition(double left, double top, double tolerance = 10)
    {
        if (_chartObjects == null || Count == 0)
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    double objLeft = chartObject.Left;
                    double objTop = chartObject.Top;

                    if (Math.Abs(objLeft - left) <= tolerance && Math.Abs(objTop - top) <= tolerance)
                    {
                        result.Add(chartObject);
                    }
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 根据大小查找图表对象
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    public IExcelChartObject[] FindBySize(double width, double height, double tolerance = 10)
    {
        if (_chartObjects == null || Count == 0)
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    double objWidth = chartObject.Width;
                    double objHeight = chartObject.Height;

                    if (Math.Abs(objWidth - width) <= tolerance && Math.Abs(objHeight - height) <= tolerance)
                    {
                        result.Add(chartObject);
                    }
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取指定区域内的所有图表对象
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的图表对象数组</returns>
    public IExcelChartObject[] GetChartsInRange(IExcelRange range)
    {
        if (_chartObjects == null || range == null || Count == 0)
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        // 注意：Excel ChartObjects集合不直接支持区域筛选
        // 这里返回所有图表对象作为示例
        for (int i = 1; i <= Count; i++)
        {
            var chartObject = this[i];
            if (chartObject != null)
                result.Add(chartObject);
        }
        return result.ToArray();
    }

    /// <summary>
    /// 获取可见的图表对象
    /// </summary>
    /// <returns>可见图表对象数组</returns>
    public IExcelChartObject[] GetVisibleCharts()
    {
        if (_chartObjects == null || Count == 0)
            return [];

        var result = new List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null && chartObject.IsVisible)
                {
                    result.Add(chartObject);
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }
        return result.ToArray();
    }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图表对象
    /// </summary>
    public void Clear()
    {
        if (_chartObjects == null) return;

        try
        {
            // 从后往前删除，避免索引变化问题
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    ((MsExcel.ChartObject)_chartObjects.Item(i)).Delete();
                }
                catch
                {
                    // 忽略删除过程中的异常
                }
            }
        }
        catch
        {
            // 忽略清空过程中的异常
        }
    }

    /// <summary>
    /// 删除指定索引的图表对象
    /// </summary>
    /// <param name="index">要删除的图表对象索引</param>
    public void Delete(int index)
    {
        if (_chartObjects == null || index < 1 || index > Count)
            return;

        try
        {
            ((MsExcel.ChartObject)_chartObjects.Item(index)).Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 删除指定的图表对象
    /// </summary>
    /// <param name="chartObject">要删除的图表对象</param>
    public void Delete(IExcelChartObject chartObject)
    {
        if (_chartObjects == null || chartObject == null)
            return;

        try
        {
            chartObject.Delete();
        }
        catch
        {
            // 忽略删除过程中的异常
        }
    }

    /// <summary>
    /// 批量删除图表对象
    /// </summary>
    /// <param name="indices">要删除的图表对象索引数组</param>
    public void DeleteRange(int[] indices)
    {
        if (_chartObjects == null || indices == null || indices.Length == 0)
            return;

        // 按降序排列索引，避免删除时索引变化
        Array.Sort(indices, (a, b) => b.CompareTo(a));

        foreach (int index in indices)
        {
            Delete(index);
        }
    }

    /// <summary>
    /// 选择所有图表对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void SelectAll(bool replace = true)
    {
        if (_chartObjects == null || Count == 0)
            return;

        try
        {
            // 选择所有图表对象
            object[] chartObjectsArray = new object[Count];
            for (int i = 1; i <= Count; i++)
            {
                chartObjectsArray[i - 1] = _chartObjects.Item(i);
            }
            // 注意：Excel中没有直接选择所有ChartObjects的方法
            // 这里提供一个空实现以保持接口一致性
        }
        catch
        {
            // 忽略选择过程中的异常
        }
    }

    /// <summary>
    /// 取消选择所有图表对象
    /// </summary>
    public void DeselectAll()
    {
        // Excel中没有直接取消选择的方法
        // 这里提供一个空实现以保持接口一致性
    }



    /// <summary>
    /// 刷新图表对象显示
    /// </summary>
    public void Refresh()
    {
        // Excel ChartObjects通常会自动刷新
        // 这里提供一个空实现以保持接口一致性
    }

    #endregion

    public IEnumerator<IExcelChartObject> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

