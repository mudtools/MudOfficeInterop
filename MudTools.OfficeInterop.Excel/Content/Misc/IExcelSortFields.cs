//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel排序字段集合的接口，用于定义和操作Excel中的排序规则
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSortFields : IOfficeObject<IExcelSortFields>, IDisposable, IEnumerable<IExcelSortField>
{
    /// <summary>
    /// 获取父级排序对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取排序字段集合中的字段数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取排序字段（索引从1开始）
    /// </summary>
    /// <param name="index">排序字段索引</param>
    /// <returns>排序字段对象</returns>
    IExcelSortField this[int index] { get; }

    /// <summary>
    /// 根据名称获取排序字段
    /// </summary>
    /// <param name="name">排序字段名称</param>
    /// <returns>排序字段对象</returns>
    IExcelSortField this[string name] { get; }

    /// <summary>
    /// 添加新的排序字段
    /// </summary>
    /// <param name="key">排序键（列范围）</param>
    /// <param name="sortOn">排序依据</param>
    /// <param name="order">排序顺序</param>
    /// <param name="customOrder">自定义排序顺序</param>
    /// <param name="dataOption">数据选项</param>
    /// <returns>新创建的排序字段对象</returns>
    IExcelSortField? Add(IExcelRange key, XlSortOn sortOn = XlSortOn.xlSortOnValues,
                       XlSortOrder order = XlSortOrder.xlAscending,
                       object? customOrder = null,
                       XlSortDataOption dataOption = XlSortDataOption.xlSortNormal);

    /// <summary>
    /// 清除所有排序字段
    /// </summary>
    void Clear();


}