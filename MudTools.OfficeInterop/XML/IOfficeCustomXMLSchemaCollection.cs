//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示附加到数据流的 Microsoft.Office.Core.CustomXMLSchema 对象集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLSchemaCollection : IOfficeObject<IOfficeCustomXMLSchemaCollection>, IEnumerable<IOfficeCustomXMLSchema?>, IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core._CustomXMLSchemaCollection 对象的父对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object? Parent { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core._CustomXMLSchemaCollection 集合中的项数。只读。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 从 Microsoft.Office.Core._CustomXMLSchemaCollection 集合中获取 CustomXMLSchema 对象。只读。
    /// </summary>
    /// <param name="index">必需 Object。要返回的 CustomXMLSchema 对象的名称或索引号。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLSchema</returns>
    IOfficeCustomXMLSchema? this[int index] { get; }

    /// <summary>
    /// 从 Microsoft.Office.Core._CustomXMLSchemaCollection 集合中获取 CustomXMLSchema 对象。只读。
    /// </summary>
    /// <param name="name">必需 Object。要返回的 CustomXMLSchema 对象的名称或索引号。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLSchema</returns>
    IOfficeCustomXMLSchema? this[string name] { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core._CustomXMLSchemaCollection 对象命名空间的唯一地址标识符。只读。
    /// </summary>
    /// <param name="index">必需 Integer。CustomXMLSchema 对象的索引号。</param>
    /// <returns>String</returns>
    [MethodIndex]
    string? NamespaceURI(int index);


    /// <summary>
    /// 允许向架构集合添加一个或多个架构，然后可以将这些架构添加到数据存储中的流和架构库中。
    /// </summary>
    /// <param name="namespaceURI">可选 String。包含要添加到集合中的架构的命名空间。如果架构已存在于架构库中，方法将从那里检索它。</param>
    /// <param name="alias">可选 String。包含要添加到集合中的架构的别名。如果别名已存在于架构库中，方法可以使用此参数找到它。</param>
    /// <param name="fileName">可选 String。包含磁盘上架构的位置。如果指定了此参数，架构将被添加到集合和架构库中。</param>
    /// <param name="installForAllUsers">可选 Boolean。指定在方法向架构库添加架构时，架构库键是否应写入注册表（HKey_Local_Machine 为所有用户，HKey_Current_User 仅为当前用户）。参数默认为 False，并写入 HKey_Current_User。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLSchema</returns>
    IOfficeCustomXMLSchema? Add(string namespaceURI = "", string alias = "", string fileName = "", bool installForAllUsers = false);

    /// <summary>
    /// 向集合中添加另一个架构集合。
    /// </summary>
    /// <param name="schemaCollection">要添加的架构集合。</param>
    void AddCollection(IOfficeCustomXMLSchemaCollection schemaCollection);

    /// <summary>
    /// 指定架构集合中的架构是否有效（符合 XML 的语法规则和指定词汇的规则；结构化 XML 的标准）。
    /// </summary>
    /// <returns>Boolean</returns>
    bool? Validate();
}