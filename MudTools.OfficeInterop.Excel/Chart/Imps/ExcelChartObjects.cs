//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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

    /// <summary>
    /// 批量添加图表对象
    /// </summary>
    /// <param name="chartData">图表数据数组</param>
    /// <returns>成功添加的图表对象数量</returns>
    public int AddRange(ChartData[] chartData)
    {
        if (_chartObjects == null || chartData == null || chartData.Length == 0)
            return 0;

        int successCount = 0;
        foreach (var data in chartData)
        {
            var chartObject = Add(data.Left, data.Top, data.Width, data.Height);
            if (chartObject != null)
            {
                if (data.SourceData != null)
                {
                    chartObject.SetSourceData(data.SourceData);
                }
                if (data.ChartType != 0)
                {
                    chartObject.SetChartType(data.ChartType);
                }
                successCount++;
            }
        }
        return successCount;
    }

    /// <summary>
    /// 基于现有数据创建图表对象
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="chartType">图表类型</param>
    /// <returns>新创建的图表对象</returns>
    public IExcelChartObject CreateFromData(IExcelRange sourceData, double left, double top,
                                          double width, double height, int chartType = 0)
    {
        if (_chartObjects == null || sourceData == null)
            return null;

        try
        {
            var chartObject = Add(left, top, width, height);
            if (chartObject != null)
            {
                var excelRange = sourceData as ExcelRange;
                chartObject.SetSourceData(sourceData);

                if (chartType != 0)
                {
                    chartObject.SetChartType(chartType);
                }
            }
            return chartObject;
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
    /// 根据图表类型查找图表对象
    /// </summary>
    /// <param name="chartType">图表类型</param>
    /// <returns>匹配的图表对象数组</returns>
    public IExcelChartObject[] FindByType(int chartType)
    {
        if (_chartObjects == null || Count == 0)
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null && chartObject.ChartType == chartType)
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
            return new IExcelChartObject[0];

        var result = new System.Collections.Generic.List<IExcelChartObject>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null && chartObject.Visible)
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

    #region 排列和布局

    /// <summary>
    /// 对齐选中的图表对象
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    public void Align(int alignment)
    {
        // Excel中没有直接的对齐方法
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 分布选中的图表对象
    /// </summary>
    /// <param name="distribution">分布方式</param>
    public void Distribute(int distribution)
    {
        // Excel中没有直接的分布方法
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 统一选中图表对象的大小
    /// </summary>
    /// <param name="useWidth">是否使用宽度作为标准</param>
    public void SizeToSame(bool useWidth = true)
    {
        // Excel中没有直接的统一大小方法
        // 这里提供一个空实现以保持接口一致性
    }

    /// <summary>
    /// 按指定行列排列图表对象
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <param name="horizontalSpacing">水平间距</param>
    /// <param name="verticalSpacing">垂直间距</param>
    public void ArrangeInGrid(int rows, int columns, double horizontalSpacing = 20, double verticalSpacing = 20)
    {
        if (_chartObjects == null || Count == 0 || rows <= 0 || columns <= 0)
            return;

        try
        {
            int chartIndex = 1;
            double baseLeft = 50; // 起始左边距
            double baseTop = 50;  // 起始顶边距

            for (int row = 0; row < rows && chartIndex <= Count; row++)
            {
                for (int col = 0; col < columns && chartIndex <= Count; col++)
                {
                    try
                    {
                        var chartObject = this[chartIndex];
                        if (chartObject != null)
                        {
                            chartObject.Left = baseLeft + col * (chartObject.Width + horizontalSpacing);
                            chartObject.Top = baseTop + row * (chartObject.Height + verticalSpacing);
                        }
                        chartIndex++;
                    }
                    catch
                    {
                        // 忽略单个图表对象排列异常
                    }
                }
            }
        }
        catch
        {
            // 忽略排列过程中的异常
        }
    }

    #endregion

    #region 导出和导入



    /// <summary>
    /// 获取所有图表对象的信息
    /// </summary>
    /// <returns>图表对象信息数组</returns>
    public ChartObjectInfo[] GetAllChartInfo()
    {
        if (_chartObjects == null || Count == 0)
            return new ChartObjectInfo[0];

        var result = new System.Collections.Generic.List<ChartObjectInfo>();
        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    var info = new ChartObjectInfo
                    {
                        Index = i,
                        Name = chartObject.Name ?? $"Chart{i}",
                        Left = chartObject.Left,
                        Top = chartObject.Top,
                        Width = chartObject.Width,
                        Height = chartObject.Height,
                        Visible = chartObject.Visible,
                        ChartType = chartObject.ChartType,
                        EnableMacro = chartObject.EnableMacro,
                        IsEmbedded = chartObject.IsEmbedded
                    };
                    result.Add(info);
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

    #region 统计和分析

    /// <summary>
    /// 获取图表对象统计信息
    /// </summary>
    /// <returns>图表对象统计信息对象</returns>
    public ChartObjectStatistics GetStatistics()
    {
        var stats = new ChartObjectStatistics
        {
            TotalCount = Count,
            VisibleCount = 0,
            HiddenCount = 0,
            AverageWidth = 0,
            AverageHeight = 0,
            MaxWidth = 0,
            MaxHeight = 0,
            UniqueTypes = 0
        };

        if (_chartObjects == null || Count == 0)
            return stats;

        double totalWidth = 0;
        double totalHeight = 0;
        var types = new System.Collections.Generic.HashSet<int>();

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    if (chartObject.Visible)
                        stats.VisibleCount++;
                    else
                        stats.HiddenCount++;

                    double width = chartObject.Width;
                    double height = chartObject.Height;
                    totalWidth += width;
                    totalHeight += height;

                    if (width > stats.MaxWidth)
                        stats.MaxWidth = width;

                    if (height > stats.MaxHeight)
                        stats.MaxHeight = height;

                    types.Add(chartObject.ChartType);
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }

        stats.AverageWidth = Count > 0 ? totalWidth / Count : 0;
        stats.AverageHeight = Count > 0 ? totalHeight / Count : 0;
        stats.UniqueTypes = types.Count;

        return stats;
    }

    /// <summary>
    /// 获取图表类型统计
    /// </summary>
    /// <returns>类型统计信息数组</returns>
    public ChartTypeStatistics[] GetTypeStatistics()
    {
        if (_chartObjects == null || Count == 0)
            return new ChartTypeStatistics[0];

        var typeCount = new System.Collections.Generic.Dictionary<int, int>();

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    int type = chartObject.ChartType;

                    if (typeCount.ContainsKey(type))
                        typeCount[type]++;
                    else
                        typeCount[type] = 1;
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }

        var result = new System.Collections.Generic.List<ChartTypeStatistics>();
        int totalCount = Count;

        foreach (var kvp in typeCount)
        {
            result.Add(new ChartTypeStatistics
            {
                ChartType = kvp.Key,
                Count = kvp.Value,
                Percentage = totalCount > 0 ? (double)kvp.Value / totalCount * 100 : 0,
                TypeName = GetChartTypeName(kvp.Key)
            });
        }

        return result.ToArray();
    }

    /// <summary>
    /// 获取图表大小分布
    /// </summary>
    /// <returns>大小分布信息</returns>
    public ChartSizeDistribution GetSizeDistribution()
    {
        var distribution = new ChartSizeDistribution
        {
            SmallCharts = 0,
            MediumCharts = 0,
            LargeCharts = 0,
            ExtraLargeCharts = 0
        };

        if (_chartObjects == null || Count == 0)
            return distribution;

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    double area = chartObject.Width * chartObject.Height;

                    if (area < 30000) // 200x150
                        distribution.SmallCharts++;
                    else if (area < 200000) // 500x400
                        distribution.MediumCharts++;
                    else if (area < 480000) // 800x600
                        distribution.LargeCharts++;
                    else
                        distribution.ExtraLargeCharts++;
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }

        return distribution;
    }

    /// <summary>
    /// 获取所有图表对象的边界框
    /// </summary>
    /// <returns>边界框信息</returns>
    public BoundingBox GetBoundingBox()
    {
        var boundingBox = new BoundingBox
        {
            Left = double.MaxValue,
            Top = double.MaxValue,
            Right = double.MinValue,
            Bottom = double.MinValue
        };

        if (_chartObjects == null || Count == 0)
        {
            boundingBox.Left = 0;
            boundingBox.Top = 0;
            boundingBox.Right = 0;
            boundingBox.Bottom = 0;
            return boundingBox;
        }

        for (int i = 1; i <= Count; i++)
        {
            try
            {
                var chartObject = this[i];
                if (chartObject != null)
                {
                    double left = chartObject.Left;
                    double top = chartObject.Top;
                    double right = left + chartObject.Width;
                    double bottom = top + chartObject.Height;

                    if (left < boundingBox.Left)
                        boundingBox.Left = left;

                    if (top < boundingBox.Top)
                        boundingBox.Top = top;

                    if (right > boundingBox.Right)
                        boundingBox.Right = right;

                    if (bottom > boundingBox.Bottom)
                        boundingBox.Bottom = bottom;
                }
            }
            catch
            {
                // 忽略单个图表对象访问异常
            }
        }

        // 确保边界框有效
        if (boundingBox.Left == double.MaxValue)
            boundingBox.Left = 0;

        if (boundingBox.Top == double.MaxValue)
            boundingBox.Top = 0;

        if (boundingBox.Right == double.MinValue)
            boundingBox.Right = 0;

        if (boundingBox.Bottom == double.MinValue)
            boundingBox.Bottom = 0;

        return boundingBox;
    }

    /// <summary>
    /// 获取图表类型名称
    /// </summary>
    /// <param name="chartType">图表类型</param>
    /// <returns>类型名称</returns>
    private string GetChartTypeName(int chartType)
    {
        switch (chartType)
        {
            case -4100: return "柱形图";
            case 5: return "堆积柱形图";
            case 6: return "百分比堆积柱形图";
            case 4: return "三维簇状柱形图";
            case -4101: return "条形图";
            case 7: return "堆积条形图";
            case 8: return "百分比堆积条形图";
            case 63: return "三维条形图";
            case -4102: return "折线图";
            case 65: return "带数据标记的折线图";
            case 66: return "带平滑线的折线图";
            case -4103: return "饼图";
            case 68: return "分离型饼图";
            case 69: return "复合饼图";
            case 70: return "复合条饼图";
            case -4104: return "散点图";
            case -4105: return "面积图";
            case 72: return "堆积面积图";
            case 73: return "百分比堆积面积图";
            case -4106: return "环形图";
            case -4107: return "雷达图";
            case 82: return "气泡图";
            case 83: return "股价图";
            default: return "未知类型";
        }
    }

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

    #endregion
}

