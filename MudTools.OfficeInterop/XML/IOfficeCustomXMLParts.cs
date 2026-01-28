//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示文档中所有自定义 XML 部分的集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLParts : IOfficeObject<IOfficeCustomXMLParts, MsCore.CustomXMLParts>, IEnumerable<IOfficeCustomXMLPart?>, IDisposable
{

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中自定义 XML 部分的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定的自定义 XML 部分
    /// </summary>
    /// <param name="index">要返回的自定义 XML 部分的从零开始的索引</param>
    /// <returns>位于指定索引处的 IOfficeCustomXMLPart 对象；如果索引无效，则返回 null</returns>
    IOfficeCustomXMLPart? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的自定义 XML 部分
    /// </summary>
    /// <param name="name">要返回的自定义 XML 部分的名称</param>
    /// <returns>具有指定名称的 IOfficeCustomXMLPart 对象；如果找不到，则返回 null</returns>
    IOfficeCustomXMLPart? this[string name] { get; }

    /// <summary>
    /// 向集合中添加一个新的自定义 XML 部分
    /// </summary>
    /// <param name="XML">要添加到新自定义 XML 部分的 XML 字符串（可选）</param>
    /// <param name="SchemaCollection">要与新自定义 XML 部分关联的架构集合（可选）</param>
    /// <returns>新创建的 IOfficeCustomXMLPart 对象；如果创建失败，则返回 null</returns>
    IOfficeCustomXMLPart? Add(string XML = "", IOfficeCustomXMLSchemaCollection? SchemaCollection = null);

    /// <summary>
    /// 通过 ID 选择指定的自定义 XML 部分
    /// </summary>
    /// <param name="id">要选择的自定义 XML 部分的 ID</param>
    /// <returns>具有指定 ID 的 IOfficeCustomXMLPart 对象；如果找不到，则返回 null</returns>
    IOfficeCustomXMLPart? SelectByID(string id);

    /// <summary>
    /// 通过命名空间 URI 选择自定义 XML 部分的集合
    /// </summary>
    /// <param name="namespaceURI">要选择的自定义 XML 部分的命名空间 URI</param>
    /// <returns>具有指定命名空间 URI 的 IOfficeCustomXMLParts 集合；如果找不到，则返回 null</returns>
    IOfficeCustomXMLParts? SelectByNamespace(string namespaceURI);
}