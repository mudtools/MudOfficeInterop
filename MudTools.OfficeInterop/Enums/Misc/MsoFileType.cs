//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文件类型，主要用于文件搜索和过滤操作
/// </summary>
public enum MsoFileType
{
    /// <summary>
    /// 所有文件类型
    /// </summary>
    msoFileTypeAllFiles = 1,

    /// <summary>
    /// Office文件类型
    /// </summary>
    msoFileTypeOfficeFiles,

    /// <summary>
    /// Word文档类型
    /// </summary>
    msoFileTypeWordDocuments,

    /// <summary>
    /// Excel工作簿类型
    /// </summary>
    msoFileTypeExcelWorkbooks,

    /// <summary>
    /// PowerPoint演示文稿类型
    /// </summary>
    msoFileTypePowerPointPresentations,

    /// <summary>
    /// 装订器文件类型
    /// </summary>
    msoFileTypeBinders,

    /// <summary>
    /// 数据库文件类型
    /// </summary>
    msoFileTypeDatabases,

    /// <summary>
    /// 模板文件类型
    /// </summary>
    msoFileTypeTemplates,

    /// <summary>
    /// Outlook项目类型
    /// </summary>
    msoFileTypeOutlookItems,

    /// <summary>
    /// 邮件项目类型
    /// </summary>
    msoFileTypeMailItem,

    /// <summary>
    /// 日历项目类型
    /// </summary>
    msoFileTypeCalendarItem,

    /// <summary>
    /// 联系人项目类型
    /// </summary>
    msoFileTypeContactItem,

    /// <summary>
    /// 笔记项目类型
    /// </summary>
    msoFileTypeNoteItem,

    /// <summary>
    /// 日记项目类型
    /// </summary>
    msoFileTypeJournalItem,

    /// <summary>
    /// 任务项目类型
    /// </summary>
    msoFileTypeTaskItem,

    /// <summary>
    /// PhotoDraw文件类型
    /// </summary>
    msoFileTypePhotoDrawFiles,

    /// <summary>
    /// 数据连接文件类型
    /// </summary>
    msoFileTypeDataConnectionFiles,

    /// <summary>
    /// Publisher文件类型
    /// </summary>
    msoFileTypePublisherFiles,

    /// <summary>
    /// Project文件类型
    /// </summary>
    msoFileTypeProjectFiles,

    /// <summary>
    /// 文档影像文件类型
    /// </summary>
    msoFileTypeDocumentImagingFiles,

    /// <summary>
    /// Visio文件类型
    /// </summary>
    msoFileTypeVisioFiles,

    /// <summary>
    /// Designer文件类型
    /// </summary>
    msoFileTypeDesignerFiles,

    /// <summary>
    /// 网页文件类型
    /// </summary>
    msoFileTypeWebPages
}