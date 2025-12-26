//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 定义Word文档中脚注选项的接口，提供对脚注位置、编号样式、起始编号等属性的访问和设置功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFootnoteOptions : IOfficeObject<IWordFootnoteOptions>, IDisposable
{

    /// <summary>
    /// 获取与此脚注选项关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此脚注选项的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置脚注的位置（如页面底部或节底部）
    /// </summary>
    WdFootnoteLocation Location { get; set; }

    /// <summary>
    /// 获取或设置脚注编号的样式（如阿拉伯数字、罗马数字、字母等）
    /// </summary>
    WdNoteNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置脚注的起始编号
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置脚注的编号规则（如连续编号、每页重新编号等）
    /// </summary>
    WdNumberingRule NumberingRule { get; set; }

    /// <summary>
    /// 获取或设置用于布局的列数
    /// </summary>
    int LayoutColumns { get; set; }
}