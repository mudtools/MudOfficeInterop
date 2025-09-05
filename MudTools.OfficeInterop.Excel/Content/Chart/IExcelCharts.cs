//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Charts 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Charts 的安全访问和操作
/// </summary>
public interface IExcelCharts : IEnumerable<IExcelChart>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取图表集合中的图表数量
    /// 对应 Charts.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的图表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">图表索引（从1开始）</param>
    /// <returns>图表对象</returns>
    IExcelChart this[int index] { get; }

    /// <summary>
    /// 获取指定名称的图表对象
    /// </summary>
    /// <param name="name">图表名称</param>
    /// <returns>图表对象</returns>
    IExcelChart this[string name] { get; }

    /// <summary>
    /// 获取图表集合所在的父对象（通常是工作表或工作簿）
    /// 对应 Charts.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取图表集合所在的Application对象
    /// 对应 Charts.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向集合中添加新的图表
    /// 对应 Charts.Add 方法
    /// </summary>
    /// <returns>新创建的图表对象</returns>
    IExcelChart Add(object Before, object After, object Count);
    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图表
    /// </summary>
    /// <param name="name">图表名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的图表数组</returns>
    IExcelChart[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据类型查找图表
    /// </summary>
    /// <param name="chartType">图表类型</param>
    /// <returns>匹配的图表数组</returns>
    IExcelChart[] FindByType(MsoChartType chartType);

    /// <summary>
    /// 获取受保护的图表
    /// </summary>
    /// <returns>受保护图表数组</returns>
    IExcelChart[] GetProtectedCharts();

    /// <summary>
    /// 获取未受保护的图表
    /// </summary>
    /// <returns>未受保护图表数组</returns>
    IExcelChart[] GetUnprotectedCharts();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图表
    /// 对应 Charts.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的图表
    /// </summary>
    /// <param name="index">要删除的图表索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的图表
    /// </summary>
    /// <param name="name">要删除的图表名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的图表对象
    /// </summary>
    /// <param name="chart">要删除的图表对象</param>
    void Delete(IExcelChart chart);

    /// <summary>
    /// 批量删除图表
    /// </summary>
    /// <param name="indices">要删除的图表索引数组</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 选择所有图表
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 刷新图表显示
    /// </summary>
    void Refresh();

    #endregion   

    #region 导出和导入

    /// <summary>
    /// 导出所有图表到文件夹
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="format">图片格式</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的图表数量</returns>
    int ExportToFolder(string folderPath, string format = "png", string prefix = "chart_");

    /// <summary>
    /// 获取所有图表的字节数组
    /// </summary>
    /// <returns>图表字节数组</returns>
    byte[][] GetAllChartBytes();

    #endregion


}
