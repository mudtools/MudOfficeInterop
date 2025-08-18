//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ChartObjects 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartObjects 的安全访问和操作
/// </summary>
public interface IExcelChartObjects : IEnumerable<IExcelChartObject>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取图表对象集合中的图表数量
    /// 对应 ChartObjects.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的图表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">图表对象索引（从1开始）</param>
    /// <returns>图表对象</returns>
    IExcelChartObject this[int index] { get; }

    /// <summary>
    /// 获取指定名称的图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>图表对象</returns>
    IExcelChartObject this[string name] { get; }

    /// <summary>
    /// 获取图表对象集合所在的父对象（通常是工作表）
    /// 对应 ChartObjects.Parent 属性
    /// </summary>
    object Parent { get; }

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
    IExcelChartObject Add(double left, double top, double width, double height);

    /// <summary>
    /// 批量添加图表对象
    /// </summary>
    /// <param name="chartData">图表数据数组</param>
    /// <returns>成功添加的图表对象数量</returns>
    int AddRange(ChartData[] chartData);

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
    IExcelChartObject CreateFromData(IExcelRange sourceData, double left, double top,
                                   double width, double height, int chartType = 0);

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindByName(string name);

    /// <summary>
    /// 根据位置查找图表对象
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindByPosition(double left, double top, double tolerance = 10);

    /// <summary>
    /// 根据大小查找图表对象
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindBySize(double width, double height, double tolerance = 10);

    /// <summary>
    /// 根据图表类型查找图表对象
    /// </summary>
    /// <param name="chartType">图表类型</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindByType(int chartType);

    /// <summary>
    /// 获取指定区域内的所有图表对象
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的图表对象数组</returns>
    IExcelChartObject[] GetChartsInRange(IExcelRange range);

    /// <summary>
    /// 获取可见的图表对象
    /// </summary>
    /// <returns>可见图表对象数组</returns>
    IExcelChartObject[] GetVisibleCharts();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图表对象
    /// 对应 ChartObjects.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的图表对象
    /// </summary>
    /// <param name="index">要删除的图表对象索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的图表对象
    /// </summary>
    /// <param name="chartObject">要删除的图表对象</param>
    void Delete(IExcelChartObject chartObject);

    /// <summary>
    /// 批量删除图表对象
    /// </summary>
    /// <param name="indices">要删除的图表对象索引数组</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 选择所有图表对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void SelectAll(bool replace = true);

    /// <summary>
    /// 取消选择所有图表对象
    /// </summary>
    void DeselectAll();

    /// <summary>
    /// 刷新图表对象显示
    /// </summary>
    void Refresh();

    #endregion

    #region 排列和布局

    /// <summary>
    /// 对齐选中的图表对象
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    void Align(int alignment);

    /// <summary>
    /// 分布选中的图表对象
    /// </summary>
    /// <param name="distribution">分布方式</param>
    void Distribute(int distribution);

    /// <summary>
    /// 统一选中图表对象的大小
    /// </summary>
    /// <param name="useWidth">是否使用宽度作为标准</param>
    void SizeToSame(bool useWidth = true);

    /// <summary>
    /// 按指定行列排列图表对象
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <param name="horizontalSpacing">水平间距</param>
    /// <param name="verticalSpacing">垂直间距</param>
    void ArrangeInGrid(int rows, int columns, double horizontalSpacing = 20, double verticalSpacing = 20);

    #endregion

    #region 导出和导入
    /// <summary>
    /// 获取所有图表对象的信息
    /// </summary>
    /// <returns>图表对象信息数组</returns>
    ChartObjectInfo[] GetAllChartInfo();
    #endregion

    #region 统计和分析

    /// <summary>
    /// 获取图表对象统计信息
    /// </summary>
    /// <returns>图表对象统计信息对象</returns>
    ChartObjectStatistics GetStatistics();

    /// <summary>
    /// 获取图表类型统计
    /// </summary>
    /// <returns>类型统计信息数组</returns>
    ChartTypeStatistics[] GetTypeStatistics();

    /// <summary>
    /// 获取图表大小分布
    /// </summary>
    /// <returns>大小分布信息</returns>
    ChartSizeDistribution GetSizeDistribution();

    /// <summary>
    /// 获取所有图表对象的边界框
    /// </summary>
    /// <returns>边界框信息</returns>
    BoundingBox GetBoundingBox();

    #endregion
}

/// <summary>
/// 图表数据结构
/// </summary>
public class ChartData
{
    /// <summary>
    /// 左边距
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// 顶边距
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// 宽度
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// 数据源区域
    /// </summary>
    public IExcelRange SourceData { get; set; }

    /// <summary>
    /// 图表类型
    /// </summary>
    public int ChartType { get; set; }
}

/// <summary>
/// 图表对象信息结构
/// </summary>
public class ChartObjectInfo
{
    /// <summary>
    /// 图表对象索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 图表对象名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 左边距
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// 顶边距
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// 宽度
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool Visible { get; set; }

    /// <summary>
    /// 图表类型
    /// </summary>
    public int ChartType { get; set; }

    /// <summary>
    /// 是否启用宏
    /// </summary>
    public bool EnableMacro { get; set; }

    /// <summary>
    /// 是否为嵌入式图表
    /// </summary>
    public bool IsEmbedded { get; set; }
}

/// <summary>
/// 图表对象统计信息结构
/// </summary>
public class ChartObjectStatistics
{
    /// <summary>
    /// 总图表对象数
    /// </summary>
    public int TotalCount { get; set; }

    /// <summary>
    /// 可见图表对象数
    /// </summary>
    public int VisibleCount { get; set; }

    /// <summary>
    /// 隐藏图表对象数
    /// </summary>
    public int HiddenCount { get; set; }

    /// <summary>
    /// 平均宽度
    /// </summary>
    public double AverageWidth { get; set; }

    /// <summary>
    /// 平均高度
    /// </summary>
    public double AverageHeight { get; set; }

    /// <summary>
    /// 最大宽度
    /// </summary>
    public double MaxWidth { get; set; }

    /// <summary>
    /// 最大高度
    /// </summary>
    public double MaxHeight { get; set; }

    /// <summary>
    /// 唯一类型数
    /// </summary>
    public int UniqueTypes { get; set; }
}

/// <summary>
/// 图表类型统计信息结构
/// </summary>
public class ChartTypeStatistics
{
    /// <summary>
    /// 图表类型
    /// </summary>
    public int ChartType { get; set; }

    /// <summary>
    /// 图表对象数量
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    /// 占比
    /// </summary>
    public double Percentage { get; set; }

    /// <summary>
    /// 类型名称
    /// </summary>
    public string TypeName { get; set; }
}

/// <summary>
/// 图表大小分布信息结构
/// </summary>
public class ChartSizeDistribution
{
    /// <summary>
    /// 小图表数量（小于200x150像素）
    /// </summary>
    public int SmallCharts { get; set; }

    /// <summary>
    /// 中等图表数量（200x150到500x400像素）
    /// </summary>
    public int MediumCharts { get; set; }

    /// <summary>
    /// 大图表数量（500x400到800x600像素）
    /// </summary>
    public int LargeCharts { get; set; }

    /// <summary>
    /// 超大图表数量（大于800x600像素）
    /// </summary>
    public int ExtraLargeCharts { get; set; }
}
