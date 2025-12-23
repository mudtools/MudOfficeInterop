//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 定义Word文本检索模式的接口，用于配置如何从Word文档中检索文本内容。
/// 此接口允许设置视图类型、是否包含隐藏文本和域代码等选项，以控制文本检索的行为。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTextRetrievalMode : IDisposable
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
    /// 获取或设置 Word 文档的视图类型，用于指定在检索文本时使用的视图模式。
    /// </summary>
    WdViewType ViewType { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示在检索文本时是否包含隐藏文本。
    /// </summary>
    bool IncludeHiddenText { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示在检索文本时是否包含域代码。
    /// </summary>
    bool IncludeFieldCodes { get; set; }

    /// <summary>
    /// 获取此文本检索模式的副本。
    /// </summary>
    IWordTextRetrievalMode Duplicate { get; }

}
