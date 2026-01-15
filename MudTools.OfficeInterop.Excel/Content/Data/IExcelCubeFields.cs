//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 数据透视表中的多维数据集字段集合的接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCubeFields : IOfficeObject<IExcelCubeFields, MsExcel.CubeFields>, IEnumerable<IExcelCubeField>, IDisposable
{
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
    /// 获取集合中 Cube 字段的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的 Cube 字段
    /// </summary>
    /// <param name="index">要获取的字段的从零开始的索引</param>
    /// <returns>指定索引处的 Cube 字段</returns>
    IExcelCubeField? this[int index] { get; }

    /// <summary>
    /// 通过名称获取集合中的 Cube 字段
    /// </summary>
    /// <param name="name">要获取的字段的名称</param>
    /// <returns>具有指定名称的 Cube 字段</returns>
    IExcelCubeField? this[string name] { get; }

    /// <summary>
    /// 向集合中添加新的命名集字段
    /// </summary>
    /// <param name="name">命名集的名称</param>
    /// <param name="caption">字段的标题</param>
    /// <returns>新创建的 Cube 字段</returns>
    IExcelCubeField? AddSet(string name, string caption);

    /// <summary>
    /// 获取度量值字段
    /// </summary>
    /// <param name="AttributeHierarchy">属性层次结构</param>
    /// <param name="function">合并计算函数</param>
    /// <param name="caption">字段标题，可选参数</param>
    /// <returns>度量值 Cube 字段</returns>
    IExcelCubeField? GetMeasure(object AttributeHierarchy, XlConsolidationFunction function, string? caption = null);

}