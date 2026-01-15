//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel数据透视表中的计算项集合接口
/// 该接口封装了对Excel数据透视表计算项集合的操作方法和属性
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCalculatedItems : IOfficeObject<IExcelCalculatedItems, MsExcel.CalculatedItems>, IEnumerable<IExcelPivotItem?>, IDisposable
{

    /// <summary>
    /// 获取计算项集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取计算项集合所属的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 通过索引获取指定位置的计算项
    /// </summary>
    /// <param name="index">计算项在集合中的索引位置</param>
    /// <returns>指定索引位置的计算项对象</returns>
    IExcelPivotItem? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的计算项
    /// </summary>
    /// <param name="name">计算项的名称</param>
    /// <returns>具有指定名称的计算项对象</returns>
    IExcelPivotItem? this[string name] { get; }

    /// <summary>
    /// 向计算项集合中添加新的计算项
    /// </summary>
    /// <param name="name">新计算项的名称</param>
    /// <param name="formula">新计算项的公式</param>
    /// <param name="useStandardFormula">是否使用标准公式格式，如果为null则使用默认设置</param>
    /// <returns>新添加的计算项对象</returns>
    IExcelPivotItem? Add(string name, string formula, bool? useStandardFormula);

}