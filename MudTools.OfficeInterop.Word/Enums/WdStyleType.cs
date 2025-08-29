namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中样式类型的枚举
/// </summary>
public enum WdStyleType
{
    /// <summary>
    /// 段落样式类型
    /// </summary>
    wdStyleTypeParagraph = 1,
    /// <summary>
    /// 字符样式类型
    /// </summary>
    wdStyleTypeCharacter,
    /// <summary>
    /// 表格样式类型
    /// </summary>
    wdStyleTypeTable,
    /// <summary>
    /// 列表样式类型
    /// </summary>
    wdStyleTypeList,
    /// <summary>
    /// 仅段落样式类型
    /// </summary>
    wdStyleTypeParagraphOnly,
    /// <summary>
    /// 链接样式类型
    /// </summary>
    wdStyleTypeLinked
}