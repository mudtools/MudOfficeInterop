//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Word 应用程序中的任务窗格类型
/// </summary>
public enum WdTaskPanes
{
    /// <summary>
    /// 格式设置任务窗格
    /// </summary>
    wdTaskPaneFormatting,

    /// <summary>
    /// 显示格式设置任务窗格
    /// </summary>
    wdTaskPaneRevealFormatting,

    /// <summary>
    /// 邮件合并任务窗格
    /// </summary>
    wdTaskPaneMailMerge,

    /// <summary>
    /// 翻译任务窗格
    /// </summary>
    wdTaskPaneTranslate,

    /// <summary>
    /// 搜索任务窗格
    /// </summary>
    wdTaskPaneSearch,

    /// <summary>
    /// XML 结构任务窗格
    /// </summary>
    wdTaskPaneXMLStructure,

    /// <summary>
    /// 文档保护任务窗格
    /// </summary>
    wdTaskPaneDocumentProtection,

    /// <summary>
    /// 文档操作任务窗格
    /// </summary>
    wdTaskPaneDocumentActions,

    /// <summary>
    /// 共享工作区任务窗格
    /// </summary>
    wdTaskPaneSharedWorkspace,

    /// <summary>
    /// 帮助任务窗格
    /// </summary>
    wdTaskPaneHelp,

    /// <summary>
    /// 研究任务窗格
    /// </summary>
    wdTaskPaneResearch,

    /// <summary>
    /// 传真服务任务窗格
    /// </summary>
    wdTaskPaneFaxService,

    /// <summary>
    /// XML 文档任务窗格
    /// </summary>
    wdTaskPaneXMLDocument,

    /// <summary>
    /// 文档更新任务窗格
    /// </summary>
    wdTaskPaneDocumentUpdates,

    /// <summary>
    /// 签名任务窗格
    /// </summary>
    wdTaskPaneSignature,

    /// <summary>
    /// 样式检查器任务窗格
    /// </summary>
    wdTaskPaneStyleInspector,

    /// <summary>
    /// 文档管理任务窗格
    /// </summary>
    wdTaskPaneDocumentManagement,

    /// <summary>
    /// 应用样式任务窗格
    /// </summary>
    wdTaskPaneApplyStyles,

    /// <summary>
    /// 导航任务窗格
    /// </summary>
    wdTaskPaneNav,

    /// <summary>
    /// 选择任务窗格
    /// </summary>
    wdTaskPaneSelection,

    /// <summary>
    /// 校对任务窗格
    /// </summary>
    wdTaskPaneProofing,

    /// <summary>
    /// XML 映射任务窗格
    /// </summary>
    wdTaskPaneXMLMapping,

    /// <summary>
    /// 修订窗格
    /// </summary>
    wdTaskPaneRevPaneFlex,

    /// <summary>
    /// 同义词库任务窗格
    /// </summary>
    wdTaskPaneThesaurus
}