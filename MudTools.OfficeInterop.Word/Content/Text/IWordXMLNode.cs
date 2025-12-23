//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的XML节点，提供对Word XML结构中节点的访问和操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordXMLNode : IDisposable
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
    /// 获取节点的基本名称（不带命名空间前缀的名称）
    /// </summary>
    string BaseName { get; }

    /// <summary>
    /// 获取节点的命名空间URI
    /// </summary>
    string NamespaceURI { get; }

    /// <summary>
    /// 获取节点的XML表示
    /// </summary>
    string XML { get; }

    /// <summary>
    /// 指示节点是否包含子节点
    /// </summary>
    bool HasChildNodes { get; }

    /// <summary>
    /// 获取与节点关联的Word范围
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取下一个同级节点
    /// </summary>
    IWordXMLNode NextSibling { get; }

    /// <summary>
    /// 获取上一个同级节点
    /// </summary>
    IWordXMLNode PreviousSibling { get; }

    /// <summary>
    /// 获取父节点
    /// </summary>
    IWordXMLNode ParentNode { get; }

    /// <summary>
    /// 获取第一个子节点
    /// </summary>
    IWordXMLNode FirstChild { get; }

    /// <summary>
    /// 获取最后一个子节点
    /// </summary>
    IWordXMLNode LastChild { get; }

    /// <summary>
    /// 获取拥有此节点的文档
    /// </summary>
    IWordDocument OwnerDocument { get; }

    /// <summary>
    /// 获取子节点集合
    /// </summary>
    IWordXMLNodes ChildNodes { get; }

    /// <summary>
    /// 获取节点的属性集合
    /// </summary>
    IWordXMLNodes Attributes { get; }

    /// <summary>
    /// 获取子节点建议集合，用于智能标记和内容控件
    /// </summary>
    IWordXMLChildNodeSuggestions ChildNodeSuggestions { get; }

    /// <summary>
    /// 获取与节点关联的智能标记
    /// </summary>
    IWordSmartTag SmartTag { get; }

    /// <summary>
    /// 获取或设置节点的文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置占位符文本
    /// </summary>
    string PlaceholderText { get; set; }

    /// <summary>
    /// 获取或设置节点的值
    /// </summary>
    string NodeValue { get; set; }

    /// <summary>
    /// 获取验证错误文本
    /// </summary>
    string ValidationErrorText { get; }


    /// <summary>
    /// 获取节点的XML节点类型
    /// </summary>
    WdXMLNodeType NodeType { get; }

    /// <summary>
    /// 获取节点的XML层级
    /// </summary>
    WdXMLNodeLevel Level { get; }

    /// <summary>
    /// 获取节点的验证状态
    /// </summary>
    WdXMLValidationStatus ValidationStatus { get; }

    /// <summary>
    /// 获取节点的Word Open XML格式内容
    /// </summary>
    string WordOpenXML { get; }


    /// <summary>
    /// 根据XPath表达式选择单个子节点
    /// </summary>
    /// <param name="XPath">要匹配的XPath表达式</param>
    /// <param name="prefixMapping">命名空间前缀映射</param>
    /// <param name="fastSearchSkippingTextNodes">是否跳过文本节点进行快速搜索</param>
    /// <returns>匹配的XML节点，如果没有匹配项则返回null</returns>
    IWordXMLNode? SelectSingleNode(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    /// <summary>
    /// 根据XPath表达式选择多个子节点
    /// </summary>
    /// <param name="XPath">要匹配的XPath表达式</param>
    /// <param name="prefixMapping">命名空间前缀映射</param>
    /// <param name="fastSearchSkippingTextNodes">是否跳过文本节点进行快速搜索</param>
    /// <returns>匹配的XML节点集合，如果没有匹配项则返回null</returns>
    IWordXMLNodes? SelectNodes(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    /// <summary>
    /// 设置节点的验证错误状态
    /// </summary>
    /// <param name="status">验证状态</param>
    /// <param name="errorText">错误文本</param>
    /// <param name="clearedAutomatically">是否自动清除错误</param>
    void SetValidationError(WdXMLValidationStatus status, string? errorText = null, bool clearedAutomatically = true);


    /// <summary>
    /// 删除当前节点
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制当前节点到剪贴板
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切当前节点到剪贴板
    /// </summary>
    void Cut();

    /// <summary>
    /// 验证当前节点
    /// </summary>
    void Validate();

    /// <summary>
    /// 移除指定的子元素
    /// </summary>
    /// <param name="ChildElement">要移除的子元素</param>
    void RemoveChild(IWordXMLNode ChildElement);

}