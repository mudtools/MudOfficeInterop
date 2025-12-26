//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的XML架构引用，用于处理与XML架构相关的操作。
/// 此接口提供了访问和管理Word文档中XML架构引用的功能，包括获取架构信息、删除和重新加载架构引用等操作。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordXMLSchemaReference : IOfficeObject<IWordXMLSchemaReference>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取命名空间 URI。
    /// </summary>
    string NamespaceURI { get; }

    /// <summary>
    /// 获取架构引用的位置。
    /// </summary>
    string Location { get; }

    /// <summary>
    /// 删除此 XML 架构引用。
    /// </summary>
    void Delete();

    /// <summary>
    /// 重新加载此 XML 架构引用。
    /// </summary>
    void Reload();
}