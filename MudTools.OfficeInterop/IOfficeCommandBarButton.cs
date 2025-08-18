//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office命令栏中的按钮控件接口
/// </summary>
/// <remarks>
/// 该接口继承自 <see cref="IOfficeCommandBarControl"/>，提供了按钮特有的属性和方法。
/// </remarks>
public interface IOfficeCommandBarButton : IOfficeCommandBarControl
{
    int Priority { get; set; }

    /// <summary>
    /// 获取或设置按钮的超链接地址
    /// </summary>
    MsoCommandBarButtonHyperlinkType HyperlinkType { get; set; }

    /// <summary>
    /// 获取或设置按钮的图标
    /// </summary>
    int FaceId { get; set; }

    /// <summary>
    /// 获取或设置按钮的快捷键文本
    /// </summary>
    string ShortcutText { get; set; }

    /// <summary>
    /// 获取或设置按钮的状态
    /// </summary>
    MsoButtonState State { get; set; }


    MsoButtonStyle Style { get; set; }

    /// <summary>
    /// 获取或设置按钮的描述文本
    /// </summary>
    string DescriptionText { get; set; }


    /// <summary>
    /// 重置按钮图标到默认状态
    /// </summary>
    void ResetIcon();

    /// <summary>
    /// 获取按钮是否为内置按钮
    /// </summary>
    bool BuiltInFace { get; }

    /// <summary>
    /// 获取或设置按钮的图标大小
    /// </summary>
    bool IsPriorityDropped { get; }
}