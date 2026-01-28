//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft.Office.Core.CustomXMLPrefixMapping 对象的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLPrefixMappings : IOfficeObject<IOfficeCustomXMLPrefixMappings, MsCore.CustomXMLPrefixMappings>, IEnumerable<IOfficeCustomXMLPrefixMapping?>, IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLPrefixMappings 对象的父对象。只读。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLPrefixMappings 集合中的项数。只读。
    /// </summary>
    /// <returns>Integer</returns>
    int Count { get; }

    /// <summary>
    /// 从 Microsoft.Office.Core.CustomXMLPrefixMappings 集合中获取 CustomXMLPrefixMapping 对象。只读。
    /// </summary>
    /// <param name="index">要返回的 CustomXMLPrefixMapping 对象的名称或索引号。</param>
    /// <returns>IOfficeCustomXMLPrefixMapping</returns>
    IOfficeCustomXMLPrefixMapping? this[int index] { get; }

    /// <summary>
    /// 从 Microsoft.Office.Core.CustomXMLPrefixMappings 集合中获取 CustomXMLPrefixMapping 对象。只读。
    /// </summary>
    /// <param name="name">要返回的 CustomXMLPrefixMapping 对象的名称或索引号。</param>
    /// <returns>IOfficeCustomXMLPrefixMapping</returns>
    IOfficeCustomXMLPrefixMapping? this[string name] { get; }

    /// <summary>
    /// 允许添加自定义命名空间/前缀映射，以便在查询项目时使用。
    /// </summary>
    /// <param name="prefix">包含要添加到前缀映射列表的前缀。</param>
    /// <param name="namespaceURI">包含要分配给新添加前缀的命名空间。</param>
    void AddNamespace(string prefix, string namespaceURI);

    /// <summary>
    /// 允许获取与指定前缀对应的命名空间。
    /// </summary>
    /// <param name="prefix">包含前缀映射列表中的前缀。</param>
    /// <returns>String</returns>
    string? LookupNamespace(string prefix);

    /// <summary>
    /// 允许获取与指定命名空间对应的前缀。
    /// </summary>
    /// <param name="namespaceURI">包含命名空间 URI。</param>
    /// <returns>String</returns>
    string? LookupPrefix(string namespaceURI);
}