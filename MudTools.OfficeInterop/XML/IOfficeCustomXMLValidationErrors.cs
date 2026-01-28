//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft.Office.Core.CustomXMLValidationError 对象的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLValidationErrors : IOfficeObject<IOfficeCustomXMLValidationErrors, MsCore.CustomXMLValidationErrors>, IEnumerable<IOfficeCustomXMLValidationError?>, IDisposable
{
    /// <summary>
    /// 获取一个 32 位整数，指示创建 Microsoft.Office.Core.CustomXMLValidationErrors 对象的应用程序。只读。
    /// </summary>
    /// <returns>Integer</returns>
    object? Parent { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLValidationErrors 集合中的项数。只读。
    /// </summary>
    /// <returns>Integer</returns>
    int Count { get; }

    /// <summary>
    /// 从 Microsoft.Office.Core.CustomXMLValidationErrors 集合中获取 CustomXMLValidationError 对象。只读。
    /// </summary>
    /// <param name="index">要返回的 CustomXMLValidationError 对象的名称或索引号。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLValidationError</returns>
    IOfficeCustomXMLValidationError? this[int index] { get; }

    /// <summary>
    /// 向 Microsoft.Office.Core.CustomXMLValidationErrors 集合中添加包含 XML 验证错误的 CustomXMLValidationError 对象。
    /// </summary>
    /// <param name="node">表示发生错误的节点。</param>
    /// <param name="errorName">包含错误名称。</param>
    /// <param name="errorText">包含描述性错误文本。</param>
    /// <param name="clearedOnUpdate">指定当 XML 被纠正和更新时，是否从 Microsoft.Office.Core.CustomXMLValidationErrors 集合中清除错误。</param>
    void Add(IOfficeCustomXMLNode node, string errorName, string errorText = "", bool clearedOnUpdate = true);
}