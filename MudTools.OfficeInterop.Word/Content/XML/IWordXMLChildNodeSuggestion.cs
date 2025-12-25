//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中 XML 子节点建议的接口，用于操作 Word 文档中的 XML 结构建议
/// 该接口封装了 COM 对象，提供对 Word XML 节点建议功能的访问和操作能力
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordXMLChildNodeSuggestion : IDisposable
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
    /// 获取此 XML 节点的基本名称（不含命名空间前缀的名称）。
    /// </summary>
    string BaseName { get; }

    /// <summary>
    /// 获取此 XML 节点的命名空间 URI。
    /// </summary>
    string NamespaceURI { get; }

    /// <summary>
    /// 获取与此 XML 节点关联的 XML 架构引用。
    /// </summary>
    IWordXMLSchemaReference? XMLSchemaReference { get; }

    /// <summary>
    /// 将此 XML 节点插入到文档的指定范围中。
    /// </summary>
    /// <param name="range">要插入节点的文档范围，可以为 null 使用默认位置</param>
    /// <returns>表示插入的 XML 节点的对象</returns>
    IWordXMLNode? Insert(IWordRange? range);
}