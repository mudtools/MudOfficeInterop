//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的智能标签对象，提供对 Word 智能标签的访问和操作功能
/// 智能标签是 Word 中识别的术语，如人名、地点、日期等，可以与操作相关联
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSmartTag : IOfficeObject<IWordSmartTag, MsWord.SmartTag>, IDisposable
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
    /// 获取智能标签的标识名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取智能标签的 XML 表示形式
    /// </summary>
    string XML { get; }

    /// <summary>
    /// 获取表示智能标签在文档中位置的范围对象
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取与智能标签关联的自定义属性集合
    /// </summary>
    IWordCustomProperties Properties { get; }

    /// <summary>
    /// 获取与智能标签关联的可用操作集合
    /// </summary>
    IWordSmartTagActions SmartTagActions { get; }

    /// <summary>
    /// 获取表示智能标签的 XML 节点对象
    /// </summary>
    IWordXMLNode XMLNode { get; }

    /// <summary>
    /// 获取用于下载智能标签相关操作的 URL
    /// </summary>
    string DownloadURL { get; }

    /// <summary>
    /// 在文档中选择此智能标签
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此智能标签
    /// </summary>
    void Delete();
}