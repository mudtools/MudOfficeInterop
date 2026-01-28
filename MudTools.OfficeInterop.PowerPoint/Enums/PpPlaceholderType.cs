//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 占位符类型枚举
/// </summary>
public enum PpPlaceholderType
{
    /// <summary>
    /// 混合占位符类型
    /// </summary>
    ppPlaceholderMixed = -2,

    /// <summary>
    /// 标题占位符
    /// </summary>
    ppPlaceholderTitle = 1,

    /// <summary>
    /// 内容占位符
    /// </summary>
    ppPlaceholderBody = 2,

    /// <summary>
    /// 中心标题占位符
    /// </summary>
    ppPlaceholderCenterTitle = 3,

    /// <summary>
    /// 剪题占位符
    /// </summary>
    ppPlaceholderSubtitle = 4,

    /// <summary>
    /// 垂直内容占位符
    /// </summary>
    ppPlaceholderVerticalBody = 5,

    /// <summary>
    /// 垂直标题占位符
    /// </summary>
    ppPlaceholderVerticalTitle = 6,

    /// <summary>
    /// 垂直对象占位符
    /// </summary>
    ppPlaceholderVerticalObject = 7,

    /// <summary>
    /// 对象占位符
    /// </summary>
    ppPlaceholderObject = 8,

    /// <summary>
    /// 图表占位符
    /// </summary>
    ppPlaceholderChart = 9,

    /// <summary>
    /// 位图占位符
    /// </summary>
    ppPlaceholderBitmap = 10,

    /// <summary>
    /// 媒体剪辑占位符
    /// </summary>
    ppPlaceholderMediaClip = 11,

    /// <summary>
    /// 组织结构图占位符
    /// </summary>
    ppPlaceholderOrgChart = 12,

    /// <summary>
    /// 表格占位符
    /// </summary>
    ppPlaceholderTable = 13,

    /// <summary>
    /// 幻灯片编号占位符
    /// </summary>
    ppPlaceholderSlideNumber = 14,

    /// <summary>
    /// 页眉占位符
    /// </summary>
    ppPlaceholderHeader = 15,

    /// <summary>
    /// 页脚占位符
    /// </summary>
    ppPlaceholderFooter = 16,

    /// <summary>
    /// 日期占位符
    /// </summary>
    ppPlaceholderDate = 17,

    /// <summary>
    /// 垂直文本占位符
    /// </summary>
    ppPlaceholderVerticalText = 18,

    /// <summary>
    /// 剪贴画占位符
    /// </summary>
    ppPlaceholderClipArt = 19,

    /// <summary>
    /// 文本框占位符
    /// </summary>
    ppPlaceholderTextBox = 20,

    /// <summary>
    /// 图片占位符
    /// </summary>
    ppPlaceholderPicture = 21,

    /// <summary>
    /// 引用占位符
    /// </summary>
    ppPlaceholderCitation = 22,

    /// <summary>
    /// 按钮占位符
    /// </summary>
    ppPlaceholderButton = 23
}
