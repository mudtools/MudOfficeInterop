//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在搜索文件或查询条件时使用的条件类型
/// </summary>
public enum MsoCondition
{
    /// <summary>
    /// 所有文件类型
    /// </summary>
    msoConditionFileTypeAllFiles = 1,
    
    /// <summary>
    /// Office文件类型
    /// </summary>
    msoConditionFileTypeOfficeFiles,
    
    /// <summary>
    /// Word文档类型
    /// </summary>
    msoConditionFileTypeWordDocuments,
    
    /// <summary>
    /// Excel工作簿类型
    /// </summary>
    msoConditionFileTypeExcelWorkbooks,
    
    /// <summary>
    /// PowerPoint演示文稿类型
    /// </summary>
    msoConditionFileTypePowerPointPresentations,
    
    /// <summary>
    /// 装订器文件类型
    /// </summary>
    msoConditionFileTypeBinders,
    
    /// <summary>
    /// 数据库文件类型
    /// </summary>
    msoConditionFileTypeDatabases,
    
    /// <summary>
    /// 模板文件类型
    /// </summary>
    msoConditionFileTypeTemplates,
    
    /// <summary>
    /// 包含指定内容
    /// </summary>
    msoConditionIncludes,
    
    /// <summary>
    /// 包含指定短语
    /// </summary>
    msoConditionIncludesPhrase,
    
    /// <summary>
    /// 以指定内容开头
    /// </summary>
    msoConditionBeginsWith,
    
    /// <summary>
    /// 以指定内容结尾
    /// </summary>
    msoConditionEndsWith,
    
    /// <summary>
    /// 包含彼此接近的内容
    /// </summary>
    msoConditionIncludesNearEachOther,
    
    /// <summary>
    /// 与指定内容完全匹配
    /// </summary>
    msoConditionIsExactly,
    
    /// <summary>
    /// 不匹配指定内容
    /// </summary>
    msoConditionIsNot,
    
    /// <summary>
    /// 昨天
    /// </summary>
    msoConditionYesterday,
    
    /// <summary>
    /// 今天
    /// </summary>
    msoConditionToday,
    
    /// <summary>
    /// 明天
    /// </summary>
    msoConditionTomorrow,
    
    /// <summary>
    /// 上周
    /// </summary>
    msoConditionLastWeek,
    
    /// <summary>
    /// 本周
    /// </summary>
    msoConditionThisWeek,
    
    /// <summary>
    /// 下周
    /// </summary>
    msoConditionNextWeek,
    
    /// <summary>
    /// 上月
    /// </summary>
    msoConditionLastMonth,
    
    /// <summary>
    /// 本月
    /// </summary>
    msoConditionThisMonth,
    
    /// <summary>
    /// 下月
    /// </summary>
    msoConditionNextMonth,
    
    /// <summary>
    /// 任何时候
    /// </summary>
    msoConditionAnytime,
    
    /// <summary>
    /// 在任意时间范围内
    /// </summary>
    msoConditionAnytimeBetween,
    
    /// <summary>
    /// 在指定日期
    /// </summary>
    msoConditionOn,
    
    /// <summary>
    /// 在指定日期或之后
    /// </summary>
    msoConditionOnOrAfter,
    
    /// <summary>
    /// 在指定日期或之前
    /// </summary>
    msoConditionOnOrBefore,
    
    /// <summary>
    /// 在接下来的指定时间内
    /// </summary>
    msoConditionInTheNext,
    
    /// <summary>
    /// 在过去的指定时间内
    /// </summary>
    msoConditionInTheLast,
    
    /// <summary>
    /// 等于指定值
    /// </summary>
    msoConditionEquals,
    
    /// <summary>
    /// 不等于指定值
    /// </summary>
    msoConditionDoesNotEqual,
    
    /// <summary>
    /// 在两个数字之间
    /// </summary>
    msoConditionAnyNumberBetween,
    
    /// <summary>
    /// 最多(小于等于)
    /// </summary>
    msoConditionAtMost,
    
    /// <summary>
    /// 至少(大于等于)
    /// </summary>
    msoConditionAtLeast,
    
    /// <summary>
    /// 多于(大于)
    /// </summary>
    msoConditionMoreThan,
    
    /// <summary>
    /// 少于(小于)
    /// </summary>
    msoConditionLessThan,
    
    /// <summary>
    /// 是(Yes)
    /// </summary>
    msoConditionIsYes,
    
    /// <summary>
    /// 否(No)
    /// </summary>
    msoConditionIsNo,
    
    /// <summary>
    /// 包含指定形式的内容
    /// </summary>
    msoConditionIncludesFormsOf,
    
    /// <summary>
    /// 自由文本
    /// </summary>
    msoConditionFreeText,
    
    /// <summary>
    /// Outlook项目文件类型
    /// </summary>
    msoConditionFileTypeOutlookItems,
    
    /// <summary>
    /// 邮件项目类型
    /// </summary>
    msoConditionFileTypeMailItem,
    
    /// <summary>
    /// 日历项目类型
    /// </summary>
    msoConditionFileTypeCalendarItem,
    
    /// <summary>
    /// 联系人项目类型
    /// </summary>
    msoConditionFileTypeContactItem,
    
    /// <summary>
    /// 笔记项目类型
    /// </summary>
    msoConditionFileTypeNoteItem,
    
    /// <summary>
    /// 日记项目类型
    /// </summary>
    msoConditionFileTypeJournalItem,
    
    /// <summary>
    /// 任务项目类型
    /// </summary>
    msoConditionFileTypeTaskItem,
    
    /// <summary>
    /// PhotoDraw文件类型
    /// </summary>
    msoConditionFileTypePhotoDrawFiles,
    
    /// <summary>
    /// 数据连接文件类型
    /// </summary>
    msoConditionFileTypeDataConnectionFiles,
    
    /// <summary>
    /// Publisher文件类型
    /// </summary>
    msoConditionFileTypePublisherFiles,
    
    /// <summary>
    /// Project文件类型
    /// </summary>
    msoConditionFileTypeProjectFiles,
    
    /// <summary>
    /// 文档图像文件类型
    /// </summary>
    msoConditionFileTypeDocumentImagingFiles,
    
    /// <summary>
    /// Visio文件类型
    /// </summary>
    msoConditionFileTypeVisioFiles,
    
    /// <summary>
    /// Designer文件类型
    /// </summary>
    msoConditionFileTypeDesignerFiles,
    
    /// <summary>
    /// 网页文件类型
    /// </summary>
    msoConditionFileTypeWebPages,
    
    /// <summary>
    /// 等于低优先级
    /// </summary>
    msoConditionEqualsLow,
    
    /// <summary>
    /// 等于普通优先级
    /// </summary>
    msoConditionEqualsNormal,
    
    /// <summary>
    /// 等于高优先级
    /// </summary>
    msoConditionEqualsHigh,
    
    /// <summary>
    /// 不等于低优先级
    /// </summary>
    msoConditionNotEqualToLow,
    
    /// <summary>
    /// 不等于普通优先级
    /// </summary>
    msoConditionNotEqualToNormal,
    
    /// <summary>
    /// 不等于高优先级
    /// </summary>
    msoConditionNotEqualToHigh,
    
    /// <summary>
    /// 等于未开始状态
    /// </summary>
    msoConditionEqualsNotStarted,
    
    /// <summary>
    /// 等于进行中状态
    /// </summary>
    msoConditionEqualsInProgress,
    
    /// <summary>
    /// 等于已完成状态
    /// </summary>
    msoConditionEqualsCompleted,
    
    /// <summary>
    /// 等于等待他人状态
    /// </summary>
    msoConditionEqualsWaitingForSomeoneElse,
    
    /// <summary>
    /// 等于已延期状态
    /// </summary>
    msoConditionEqualsDeferred,
    
    /// <summary>
    /// 不等于未开始状态
    /// </summary>
    msoConditionNotEqualToNotStarted,
    
    /// <summary>
    /// 不等于进行中状态
    /// </summary>
    msoConditionNotEqualToInProgress,
    
    /// <summary>
    /// 不等于已完成状态
    /// </summary>
    msoConditionNotEqualToCompleted,
    
    /// <summary>
    /// 不等于等待他人状态
    /// </summary>
    msoConditionNotEqualToWaitingForSomeoneElse,
    
    /// <summary>
    /// 不等于已延期状态
    /// </summary>
    msoConditionNotEqualToDeferred
}