//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中尾注选项的接口，用于配置和管理尾注的显示位置、编号样式等属性
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordEndnoteOptions : IOfficeObject<IWordEndnoteOptions>, IDisposable
{
    /// <summary>
    /// 获取与此尾注选项关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此尾注选项的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置尾注在文档中的位置（例如文档末尾或节末尾）
    /// </summary>
    WdEndnoteLocation Location { get; set; }

    /// <summary>
    /// 获取或设置尾注的编号样式（例如阿拉伯数字、罗马数字、字母等）
    /// </summary>
    WdNoteNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置尾注的起始编号
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置尾注的编号规则（例如连续编号、每节重新编号等）
    /// </summary>
    WdNumberingRule NumberingRule { get; set; }
}