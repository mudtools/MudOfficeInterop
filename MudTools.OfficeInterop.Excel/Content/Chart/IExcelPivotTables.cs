//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotTables 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotTables 的安全访问和操作
/// </summary>
public interface IExcelPivotTables : IEnumerable<IExcelPivotTable>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表集合中的透视表数量
    /// 对应 PivotTables.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的数据透视表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">透视表索引（从1开始）</param>
    /// <returns>数据透视表对象</returns>
    IExcelPivotTable this[int index] { get; }

    /// <summary>
    /// 获取指定名称的数据透视表对象
    /// </summary>
    /// <param name="name">透视表名称</param>
    /// <returns>数据透视表对象</returns>
    IExcelPivotTable this[string name] { get; }

    /// <summary>
    /// 获取数据透视表集合所在的父对象（通常是 Worksheet）
    /// 对应 PivotTables.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取数据透视表集合所在的Application对象
    /// 对应 PivotTables.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 查找和筛选

    /// <summary>
    /// 向数据透视表集合中添加一个新的数据透视表
    /// </summary>
    /// <param name="pivotCache">数据透视表缓存对象，包含数据源信息</param>
    /// <param name="tableDestination">数据透视表放置的目标区域</param>
    /// <param name="tableName">数据透视表的名称，可为null</param>
    /// <param name="readData">是否立即读取数据，可为null</param>
    /// <returns>新创建的数据透视表对象，如果创建失败则返回null</returns>
    IExcelPivotTable? Add(IExcelPivotCache pivotCache, IExcelRange tableDestination, string? tableName = null, bool? readData = null);
    /// <summary>
    /// 根据名称查找数据透视表
    /// </summary>
    /// <param name="name">透视表名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的透视表数组</returns>
    IExcelPivotTable[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据源数据查找数据透视表
    /// </summary>
    /// <param name="sourceData">源数据区域</param>
    /// <returns>匹配的透视表数组</returns>
    IExcelPivotTable[] FindBySourceData(IExcelRange sourceData);

    /// <summary>
    /// 获取受保护的数据透视表
    /// </summary>
    /// <returns>受保护透视表数组</returns>
    IExcelPivotTable[] GetProtectedPivotTables();

    /// <summary>
    /// 获取未受保护的数据透视表
    /// </summary>
    /// <returns>未受保护透视表数组</returns>
    IExcelPivotTable[] GetUnprotectedPivotTables();
    #endregion
}
