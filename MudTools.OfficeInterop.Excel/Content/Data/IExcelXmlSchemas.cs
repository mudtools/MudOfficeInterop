//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel XML 架构集合的接口，提供对 XML 架构集合的枚举和资源管理功能。
/// 此接口继承自 IEnumerable[IExcelXmlSchema] 和 IDisposable，支持遍历集合中的 XML 架构以及资源释放。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelXmlSchemas : IOfficeObject<IExcelXmlSchemas>, IEnumerable<IExcelXmlSchema?>, IDisposable
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
    /// 获取集合中 XML 架构定义的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引位置获取集合中的 XML 架构。
    /// </summary>
    /// <param name="index">XML 架构在集合中的从 1 开始的索引位置。</param>
    /// <returns>指定索引位置的 XML 架构。</returns>
    IExcelXmlSchema? this[int index] { get; }

    /// <summary>
    /// 通过名称获取集合中的 XML 架构。
    /// </summary>
    /// <param name="name">XML 架构的名称。</param>
    /// <returns>具有指定名称的 XML 架构。</returns>
    IExcelXmlSchema? this[string name] { get; }
}
