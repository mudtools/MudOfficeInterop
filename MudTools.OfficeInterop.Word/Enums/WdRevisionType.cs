//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的修订类型
/// </summary>
public enum WdRevisionType
{
    /// <summary>
    /// 无修订
    /// </summary>
    wdNoRevision,
    /// <summary>
    /// 插入内容修订
    /// </summary>
    wdRevisionInsert,
    /// <summary>
    /// 删除内容修订
    /// </summary>
    wdRevisionDelete,
    /// <summary>
    /// 属性修订
    /// </summary>
    wdRevisionProperty,
    /// <summary>
    /// 段落编号修订
    /// </summary>
    wdRevisionParagraphNumber,
    /// <summary>
    /// 显示字段修订
    /// </summary>
    wdRevisionDisplayField,
    /// <summary>
    /// 调和修订
    /// </summary>
    wdRevisionReconcile,
    /// <summary>
    /// 冲突修订
    /// </summary>
    wdRevisionConflict,
    /// <summary>
    /// 样式修订
    /// </summary>
    wdRevisionStyle,
    /// <summary>
    /// 替换修订
    /// </summary>
    wdRevisionReplace,
    /// <summary>
    /// 段落属性修订
    /// </summary>
    wdRevisionParagraphProperty,
    /// <summary>
    /// 表格属性修订
    /// </summary>
    wdRevisionTableProperty,
    /// <summary>
    /// 节属性修订
    /// </summary>
    wdRevisionSectionProperty,
    /// <summary>
    /// 样式定义修订
    /// </summary>
    wdRevisionStyleDefinition,
    /// <summary>
    /// 移出修订
    /// </summary>
    wdRevisionMovedFrom,
    /// <summary>
    /// 移入修订
    /// </summary>
    wdRevisionMovedTo,
    /// <summary>
    /// 单元格插入修订
    /// </summary>
    wdRevisionCellInsertion,
    /// <summary>
    /// 单元格删除修订
    /// </summary>
    wdRevisionCellDeletion,
    /// <summary>
    /// 单元格合并修订
    /// </summary>
    wdRevisionCellMerge,
    /// <summary>
    /// 单元格拆分修订
    /// </summary>
    wdRevisionCellSplit,
    /// <summary>
    /// 冲突插入修订
    /// </summary>
    wdRevisionConflictInsert,
    /// <summary>
    /// 冲突删除修订
    /// </summary>
    wdRevisionConflictDelete
}