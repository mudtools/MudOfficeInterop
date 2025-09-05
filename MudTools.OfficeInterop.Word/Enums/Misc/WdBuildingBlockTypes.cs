namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中构建基块的类型。
/// 构建基块是可重复使用的内容单元，可以快速插入到文档中。
/// </summary>
public enum WdBuildingBlockTypes
{
    /// <summary>
    /// 快速部件构建基块类型
    /// </summary>
    wdTypeQuickParts = 1,

    /// <summary>
    /// 封面页构建基块类型
    /// </summary>
    wdTypeCoverPage,

    /// <summary>
    /// 公式构建基块类型
    /// </summary>
    wdTypeEquations,

    /// <summary>
    /// 页脚构建基块类型
    /// </summary>
    wdTypeFooters,

    /// <summary>
    /// 页眉构建基块类型
    /// </summary>
    wdTypeHeaders,

    /// <summary>
    /// 页码构建基块类型
    /// </summary>
    wdTypePageNumber,

    /// <summary>
    /// 表格构建基块类型
    /// </summary>
    wdTypeTables,

    /// <summary>
    /// 水印构建基块类型
    /// </summary>
    wdTypeWatermarks,

    /// <summary>
    /// 自动图文集构建基块类型
    /// </summary>
    wdTypeAutoText,

    /// <summary>
    /// 文本框构建基块类型
    /// </summary>
    wdTypeTextBox,

    /// <summary>
    /// 顶部页码构建基块类型
    /// </summary>
    wdTypePageNumberTop,

    /// <summary>
    /// 底部页码构建基块类型
    /// </summary>
    wdTypePageNumberBottom,

    /// <summary>
    /// 页面页码构建基块类型
    /// </summary>
    wdTypePageNumberPage,

    /// <summary>
    /// 目录构建基块类型
    /// </summary>
    wdTypeTableOfContents,

    /// <summary>
    /// 自定义快速部件构建基块类型
    /// </summary>
    wdTypeCustomQuickParts,

    /// <summary>
    /// 自定义封面页构建基块类型
    /// </summary>
    wdTypeCustomCoverPage,

    /// <summary>
    /// 自定义公式构建基块类型
    /// </summary>
    wdTypeCustomEquations,

    /// <summary>
    /// 自定义页脚构建基块类型
    /// </summary>
    wdTypeCustomFooters,

    /// <summary>
    /// 自定义页眉构建基块类型
    /// </summary>
    wdTypeCustomHeaders,

    /// <summary>
    /// 自定义页码构建基块类型
    /// </summary>
    wdTypeCustomPageNumber,

    /// <summary>
    /// 自定义表格构建基块类型
    /// </summary>
    wdTypeCustomTables,

    /// <summary>
    /// 自定义水印构建基块类型
    /// </summary>
    wdTypeCustomWatermarks,

    /// <summary>
    /// 自定义自动图文集构建基块类型
    /// </summary>
    wdTypeCustomAutoText,

    /// <summary>
    /// 自定义文本框构建基块类型
    /// </summary>
    wdTypeCustomTextBox,

    /// <summary>
    /// 自定义顶部页码构建基块类型
    /// </summary>
    wdTypeCustomPageNumberTop,

    /// <summary>
    /// 自定义底部页码构建基块类型
    /// </summary>
    wdTypeCustomPageNumberBottom,

    /// <summary>
    /// 自定义页面页码构建基块类型
    /// </summary>
    wdTypeCustomPageNumberPage,

    /// <summary>
    /// 自定义目录构建基块类型
    /// </summary>
    wdTypeCustomTableOfContents,

    /// <summary>
    /// 自定义类型1构建基块
    /// </summary>
    wdTypeCustom1,

    /// <summary>
    /// 自定义类型2构建基块
    /// </summary>
    wdTypeCustom2,

    /// <summary>
    /// 自定义类型3构建基块
    /// </summary>
    wdTypeCustom3,

    /// <summary>
    /// 自定义类型4构建基块
    /// </summary>
    wdTypeCustom4,

    /// <summary>
    /// 自定义类型5构建基块
    /// </summary>
    wdTypeCustom5,

    /// <summary>
    /// 参考文献构建基块类型
    /// </summary>
    wdTypeBibliography,

    /// <summary>
    /// 自定义参考文献构建基块类型
    /// </summary>
    wdTypeCustomBibliography
}