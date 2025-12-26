//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel数据透视表更改列表的接口，提供对数据透视表中值更改的集合访问功能。
/// 继承自IEnumerable[IExcelValueChange]和IDisposable接口，支持遍历和资源释放。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPivotTableChangeList : IOfficeObject<IExcelPivotTableChangeList>, IEnumerable<IExcelValueChange>, IDisposable
{

    /// <summary>
    /// 获取集合中元素的数量。
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 通过索引获取指定位置的值更改项。
    /// </summary>
    /// <param name="index">要获取的元素从零开始的索引。</param>
    /// <returns>指定索引处的IExcelValueChange对象，如果索引无效则返回null。</returns>
    IExcelValueChange? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的值更改项。
    /// </summary>
    /// <param name="name">要获取的元素的名称。</param>
    /// <returns>具有指定名称的IExcelValueChange对象，如果未找到则返回null。</returns>
    IExcelValueChange? this[string name] { get; }


    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 向数据透视表更改列表中添加一个新的值更改项。
    /// </summary>
    /// <param name="tuple">表示元组的字符串。</param>
    /// <param name="value">要设置的数值。</param>
    /// <param name="allocationValue">分配值类型，可空，默认为null。</param>
    /// <param name="allocationMethod">分配方法，可空，默认为null。</param>
    /// <param name="allocationWeightExpression">分配权重表达式，可空，默认为null。</param>
    /// <returns>返回新添加的Excel值更改对象。</returns>
    IExcelValueChange Add(string tuple, double value,
        XlAllocationValue? allocationValue = null, XlAllocationMethod? allocationMethod = null,
        string? allocationWeightExpression = null);
}