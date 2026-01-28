//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 视图类型枚举
/// </summary>
public enum PpViewType
{
    /// <summary>
    /// 幻灯片视图
    /// </summary>
    ppViewSlide = 1,

    /// <summary>
    /// 幻灯片母版视图
    /// </summary>
    ppViewSlideMaster = 2,

    /// <summary>
    /// 备注页视图
    /// </summary>
    ppViewNotesPage = 3,

    /// <summary>
    /// 讲义母版视图
    /// </summary>
    ppViewHandoutMaster = 4,

    /// <summary>
    /// 备注母版视图
    /// </summary>
    ppViewNotesMaster = 5,

    /// <summary>
    /// 大纲视图
    /// </summary>
    ppViewOutline = 6,

    /// <summary>
    /// 幻灯片浏览视图
    /// </summary>
    ppViewSlideSorter = 7,

    /// <summary>
    /// 标题母版视图
    /// </summary>
    ppViewTitleMaster = 8,

    /// <summary>
    /// 普通视图
    /// </summary>
    ppViewNormal = 9,

    /// <summary>
    /// 打印预览视图
    /// </summary>
    ppViewPrintPreview = 10,

    /// <summary>
    /// 缩略图视图
    /// </summary>
    ppViewThumbnails = 11,

    /// <summary>
    /// 母版缩略图视图
    /// </summary>
    ppViewMasterThumbnails = 12
}
