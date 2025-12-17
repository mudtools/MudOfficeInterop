//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;



[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelCalculatedMembers : IEnumerable<IExcelCalculatedMember>, IDisposable
{
    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 根据索引获取集合中的计算成员。
    /// </summary>
    /// <param name="index">要获取的计算成员的从 1 开始的索引。</param>
    /// <returns>指定索引处的计算成员，如果不存在则返回 null。</returns>
    IExcelCalculatedMember? this[int index] { get; }

    /// <summary>
    /// 根据名称获取集合中的计算成员。
    /// </summary>
    /// <param name="name">要获取的计算成员的名称。</param>
    /// <returns>指定名称的计算成员，如果不存在则返回 null。</returns>
    IExcelCalculatedMember? this[string name] { get; }

    /// <summary>
    /// 获取集合中计算成员的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 向集合中添加新的计算成员。
    /// </summary>
    /// <param name="name">计算成员的名称。</param>
    /// <param name="formula">计算成员的公式。</param>
    /// <param name="solveOrder">求解顺序，可为空。</param>
    /// <param name="type">计算成员类型，可为空。</param>
    /// <returns>新创建的计算成员。</returns>
    IExcelCalculatedMember? Add(string name, string formula, int? solveOrder = null, XlCalculatedMemberType? type = null);

    /// <summary>
    /// 向集合中添加新的计算成员（增强版）。
    /// </summary>
    /// <param name="name">计算成员的名称。</param>
    /// <param name="formula">计算成员的公式，可为空。</param>
    /// <param name="solveOrder">求解顺序，可为空。</param>
    /// <param name="type">计算成员类型，可为空。</param>
    /// <param name="dynamic">是否为动态计算成员，可为空。</param>
    /// <param name="displayFolder">显示文件夹路径，可为空。</param>
    /// <param name="hierarchizeDistinct">是否分层区分，可为空。</param>
    /// <returns>新创建的计算成员。</returns>
    IExcelCalculatedMember? Add2(string name, string? formula = null, int? solveOrder = null, XlCalculatedMemberType? type = null, bool? dynamic = null, string? displayFolder = null, bool? hierarchizeDistinct = null);

    /// <summary>
    /// 添加计算成员到集合中。
    /// </summary>
    /// <param name="name">计算成员的名称。</param>
    /// <param name="formula">计算成员的公式，可为空。</param>
    /// <param name="solveOrder">求解顺序，可为空。</param>
    /// <param name="type">计算成员类型，可为空。</param>
    /// <param name="displayFolder">显示文件夹路径，可为空。</param>
    /// <param name="measureGroup">度量值组名称，可为空。</param>
    /// <param name="parentHierarchy">上级层次结构名称，可为空。</param>
    /// <param name="parentMember">上级成员名称，可为空。</param>
    /// <param name="numberFormat">数字格式类型，可为空。</param>
    /// <returns>新创建的计算成员。</returns>
    IExcelCalculatedMember? AddCalculatedMember(string name, string? formula = null, int? solveOrder = null, XlCalculatedMemberType? type = null, string? displayFolder = null, string? measureGroup = null, string? parentHierarchy = null, string? parentMember = null, XlCalcMemNumberFormatType? numberFormat = null);

}
