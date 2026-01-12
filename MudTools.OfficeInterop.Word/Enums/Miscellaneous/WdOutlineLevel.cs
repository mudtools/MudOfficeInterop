namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定段落的大纲级别。大纲级别用于创建文档结构，通常与标题级别相关联。
/// 数字级别1-9表示不同级别的标题，wdOutlineLevelBodyText表示正文文本。
/// </summary>
public enum WdOutlineLevel
{
    /// <summary>
    /// 一级大纲级别，通常对应文档的主要章节标题
    /// </summary>
    wdOutlineLevel1 = 1,
    /// <summary>
    /// 二级大纲级别
    /// </summary>
    wdOutlineLevel2,
    /// <summary>
    /// 三级大纲级别
    /// </summary>
    wdOutlineLevel3,
    /// <summary>
    /// 四级大纲级别
    /// </summary>
    wdOutlineLevel4,
    /// <summary>
    /// 五级大纲级别
    /// </summary>
    wdOutlineLevel5,
    /// <summary>
    /// 六级大纲级别
    /// </summary>
    wdOutlineLevel6,
    /// <summary>
    /// 七级大纲级别
    /// </summary>
    wdOutlineLevel7,
    /// <summary>
    /// 八级大纲级别
    /// </summary>
    wdOutlineLevel8,
    /// <summary>
    /// 九级大纲级别
    /// </summary>
    wdOutlineLevel9,
    /// <summary>
    /// 正文文本级别，不包含在大纲中
    /// </summary>
    wdOutlineLevelBodyText
}