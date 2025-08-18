//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotCaches 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotCaches 的安全访问和操作
/// </summary>
public interface IExcelPivotCaches : IEnumerable<IExcelPivotCache>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表缓存集合中的缓存数量
    /// 对应 PivotCaches.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的数据透视表缓存对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">缓存索引（从1开始）</param>
    /// <returns>数据透视表缓存对象</returns>
    IExcelPivotCache? this[int index] { get; }

    /// <summary>
    /// 获取数据透视表缓存集合所在的父对象（通常是 Workbook）
    /// 对应 PivotCaches.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表缓存集合所在的Application对象
    /// 对应 PivotCaches.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的数据透视表缓存
    /// 对应 PivotCaches.Create 方法
    /// </summary>
    /// <param name="sourceType">数据源类型</param>
    /// <param name="sourceData">数据源</param>
    /// <param name="version">数据透视表版本</param>
    /// <returns>新创建的数据透视表缓存对象</returns>
    IExcelPivotCache Create(int sourceType, object sourceData, object version = null);
    #endregion

    #region 查找和筛选  

    /// <summary>
    /// 根据内存使用量查找缓存 (占位符)
    /// </summary>
    /// <param name="minSize">最小内存大小 (字节)</param>
    /// <returns>匹配的缓存对象数组</returns>
    IExcelPivotCache[] FindByMemoryUsage(long minSize);
    #endregion

    #region 操作方法
    /// <summary>
    /// 刷新所有数据透视表缓存
    /// </summary>
    void Refresh();
    #endregion

    #region 导出和导入
    // PivotCaches 本身不直接导出/导入，但可以导出关联的数据或信息
    /// <summary>
    /// 导出所有缓存的元数据信息到文件夹
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="format">导出格式 (例如 "json", "xml")</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的缓存信息数量</returns>
    int ExportMetadataToFolder(string folderPath, string format = "json", string prefix = "pivotcache_");
    #endregion

    #region 高级功能  

    /// <summary>
    /// 刷新所有数据透视表缓存并更新关联的数据透视表
    /// </summary>
    void RefreshAll();
    #endregion
}
