//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 UserAccess 对象的集合，这些对象表示受保护区域的用户访问权限。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelUserAccessList : IEnumerable<IExcelUserAccess?>, IOfficeObject<IExcelUserAccessList, MsExcel.UserAccessList>, IDisposable
{

    /// <summary>
    /// 获取集合中的对象数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的 UserAccess 对象（索引器）。
    /// </summary>
    /// <param name="index">对象的名称或索引号。</param>
    /// <returns>指定索引处的 UserAccess 对象。</returns>
    IExcelUserAccess? this[int index] { get; }

    /// <summary>
    /// 获取指定索引处的 UserAccess 对象（索引器）。
    /// </summary>
    /// <param name="name">对象的名称或索引号。</param>
    /// <returns>指定索引处的 UserAccess 对象。</returns>
    IExcelUserAccess? this[string name] { get; }

    /// <summary>
    /// 添加用户访问列表。返回 UserAccess 对象。
    /// </summary>
    /// <param name="name">必需。用户访问列表的名称。</param>
    /// <param name="allowEdit">必需。布尔值。True 表示允许访问列表中的用户编辑受保护工作表中的可编辑区域。</param>
    /// <returns>新创建的 UserAccess 对象。</returns>
    IExcelUserAccess? Add(string name, bool allowEdit);

    /// <summary>
    /// 删除与工作表中受保护区域访问权限关联的所有用户。
    /// </summary>
    void DeleteAll();
}