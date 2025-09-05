//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定要返回的有关文档位置的信息类型
/// </summary>
public enum WdInformation
{
    /// <summary>
    /// 当前活动页眉或页脚中调整后的页码
    /// </summary>
    wdActiveEndAdjustedPageNumber = 1,
    
    /// <summary>
    /// 区域或选定内容结束位置的节号
    /// </summary>
    wdActiveEndSectionNumber = 2,
    
    /// <summary>
    /// 区域或选定内容结束位置的页码
    /// </summary>
    wdActiveEndPageNumber = 3,
    
    /// <summary>
    /// 文档中的总页数
    /// </summary>
    wdNumberOfPagesInDocument = 4,
    
    /// <summary>
    /// 相对于页面左边界的水平位置
    /// </summary>
    wdHorizontalPositionRelativeToPage = 5,
    
    /// <summary>
    /// 相对于页面上边界的垂直位置
    /// </summary>
    wdVerticalPositionRelativeToPage = 6,
    
    /// <summary>
    /// 相对于文本边界左边的水平位置
    /// </summary>
    wdHorizontalPositionRelativeToTextBoundary = 7,
    
    /// <summary>
    /// 相对于文本边界上边的垂直位置
    /// </summary>
    wdVerticalPositionRelativeToTextBoundary = 8,
    
    /// <summary>
    /// 范围内第一个字符的列号
    /// </summary>
    wdFirstCharacterColumnNumber = 9,
    
    /// <summary>
    /// 范围内第一个字符的行号
    /// </summary>
    wdFirstCharacterLineNumber = 10,
    
    /// <summary>
    /// 是否选定框架
    /// </summary>
    wdFrameIsSelected = 11,
    
    /// <summary>
    /// 是否在表格内
    /// </summary>
    wdWithInTable = 12,
    
    /// <summary>
    /// 范围开始行号
    /// </summary>
    wdStartOfRangeRowNumber = 13,
    
    /// <summary>
    /// 范围结束行号
    /// </summary>
    wdEndOfRangeRowNumber = 14,
    
    /// <summary>
    /// 表格中的最大行数
    /// </summary>
    wdMaximumNumberOfRows = 15,
    
    /// <summary>
    /// 范围开始列号
    /// </summary>
    wdStartOfRangeColumnNumber = 16,
    
    /// <summary>
    /// 范围结束列号
    /// </summary>
    wdEndOfRangeColumnNumber = 17,
    
    /// <summary>
    /// 表格中的最大列数
    /// </summary>
    wdMaximumNumberOfColumns = 18,
    
    /// <summary>
    /// 当前缩放比例
    /// </summary>
    wdZoomPercentage = 19,
    
    /// <summary>
    /// 当前选择模式
    /// </summary>
    wdSelectionMode = 20,
    
    /// <summary>
    /// Caps Lock键状态
    /// </summary>
    wdCapsLock = 21,
    
    /// <summary>
    /// Num Lock键状态
    /// </summary>
    wdNumLock = 22,
    
    /// <summary>
    /// 改写模式状态
    /// </summary>
    wdOverType = 23,
    
    /// <summary>
    /// 修订标记状态
    /// </summary>
    wdRevisionMarking = 24,
    
    /// <summary>
    /// 是否在脚注或尾注窗格中
    /// </summary>
    wdInFootnoteEndnotePane = 25,
    
    /// <summary>
    /// 是否在批注窗格中
    /// </summary>
    wdInCommentPane = 26,
    
    /// <summary>
    /// 是否在页眉或页脚中
    /// </summary>
    wdInHeaderFooter = 28,
    
    /// <summary>
    /// 是否在行标记末尾
    /// </summary>
    wdAtEndOfRowMarker = 31,
    
    /// <summary>
    /// 引用类型
    /// </summary>
    wdReferenceOfType = 32,
    
    /// <summary>
    /// 页眉或页脚类型
    /// </summary>
    wdHeaderFooterType = 33,
    
    /// <summary>
    /// 是否在主文档中
    /// </summary>
    wdInMasterDocument = 34,
    
    /// <summary>
    /// 是否在脚注中
    /// </summary>
    wdInFootnote = 35,
    
    /// <summary>
    /// 是否在尾注中
    /// </summary>
    wdInEndnote = 36,
    
    /// <summary>
    /// 是否在Word邮件中
    /// </summary>
    wdInWordMail = 37,
    
    /// <summary>
    /// 是否在剪贴板中
    /// </summary>
    wdInClipboard = 38,
    
    /// <summary>
    /// 是否在封面页中
    /// </summary>
    wdInCoverPage = 41,
    
    /// <summary>
    /// 是否在参考文献中
    /// </summary>
    wdInBibliography = 42,
    
    /// <summary>
    /// 是否在引用中
    /// </summary>
    wdInCitation = 43,
    
    /// <summary>
    /// 是否在域代码中
    /// </summary>
    wdInFieldCode = 44,
    
    /// <summary>
    /// 是否在域结果中
    /// </summary>
    wdInFieldResult = 45,
    
    /// <summary>
    /// 是否在内容控件中
    /// </summary>
    wdInContentControl = 46
}