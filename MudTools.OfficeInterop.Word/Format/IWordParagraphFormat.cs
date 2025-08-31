namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ParagraphFormat 的接口，用于操作 Word 文档中段落的格式设置。
/// </summary>
public interface IWordParagraphFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置段落的对齐方式。
    /// </summary>
    WdParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置首行缩进距离（单位为磅）。
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置左缩进距离（单位为磅）。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置右缩进距离（单位为磅）。
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置段落前间距（单位为磅）。
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段落后间距（单位为磅）。
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置行距规则。
    /// </summary>
    WdLineSpacing LineSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置行距值（当 LineSpacingRule 为 wdLineSpaceExactly 或 wdLineSpaceMultiple 时使用）。
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置是否禁止孤行控制（防止段落的第一行或最后一行单独出现在页面顶部或底部）。
    /// </summary>
    bool WidowControl { get; set; }

    /// <summary>
    /// 获取或设置是否保持段落在一起（分页时段落不会被分割）。
    /// </summary>
    bool KeepTogether { get; set; }

    /// <summary>
    /// 获取或设置段落是否与下一段保持在同一页面。
    /// </summary>
    bool KeepWithNext { get; set; }

    /// <summary>
    /// 获取或设置段落的制表符位置集合。
    /// </summary>
    IWordTabStops? TabStops { get; }

    /// <summary>
    /// 获取或设置段落大纲级别。
    /// </summary>
    WdOutlineLevel OutlineLevel { get; set; }

    /// <summary>
    /// 获取或设置段落的边框样式。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取或设置段落的底纹样式。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取或设置段落的文本方向。
    /// </summary>
    WdReadingOrder ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置是否启用字符网格对齐。
    /// </summary>
    bool CharacterUnitLeftIndent { get; set; }

    /// <summary>
    /// 获取或设置是否启用字符单位的首行缩进。
    /// </summary>
    bool CharacterUnitFirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置是否启用字符单位的右缩进。
    /// </summary>
    bool CharacterUnitRightIndent { get; set; }
}