//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定从剪贴板粘贴内容时的恢复格式类型
/// </summary>
public enum WdRecoveryType
{
    /// <summary>
    /// 使用默认粘贴格式
    /// </summary>
    wdPasteDefault = 0,
    
    /// <summary>
    /// 将文本粘贴到单个单元格中
    /// </summary>
    wdSingleCellText = 5,
    
    /// <summary>
    /// 将表格粘贴到单个单元格中
    /// </summary>
    wdSingleCellTable = 6,
    
    /// <summary>
    /// 继续列表编号
    /// </summary>
    wdListContinueNumbering = 7,
    
    /// <summary>
    /// 重新开始列表编号
    /// </summary>
    wdListRestartNumbering = 8,
    
    /// <summary>
    /// 将表格作为行插入
    /// </summary>
    wdTableInsertAsRows = 11,
    
    /// <summary>
    /// 追加表格
    /// </summary>
    wdTableAppendTable = 10,
    
    /// <summary>
    /// 保持表格原始格式
    /// </summary>
    wdTableOriginalFormatting = 12,
    
    /// <summary>
    /// 图片形式的图表
    /// </summary>
    wdChartPicture = 13,
    
    /// <summary>
    /// 图表对象
    /// </summary>
    wdChart = 14,
    
    /// <summary>
    /// 链接的图表
    /// </summary>
    wdChartLinked = 15,
    
    /// <summary>
    /// 保持原始格式
    /// </summary>
    wdFormatOriginalFormatting = 16,
    
    /// <summary>
    /// 使用周围格式并强调
    /// </summary>
    wdFormatSurroundingFormattingWithEmphasis = 20,
    
    /// <summary>
    /// 纯文本格式
    /// </summary>
    wdFormatPlainText = 22,
    
    /// <summary>
    /// 覆盖单元格内容
    /// </summary>
    wdTableOverwriteCells = 23,
    
    /// <summary>
    /// 合并到现有列表
    /// </summary>
    wdListCombineWithExistingList = 24,
    
    /// <summary>
    /// 不合并列表
    /// </summary>
    wdListDontMerge = 25,
    
    /// <summary>
    /// 使用目标文档样式恢复
    /// </summary>
    wdUseDestinationStylesRecovery = 19
}