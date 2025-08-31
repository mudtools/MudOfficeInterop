namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中链接类型的枚举
/// </summary>
public enum WdLinkType
{
    /// <summary>
    /// OLE对象链接
    /// </summary>
    wdLinkTypeOLE,
    /// <summary>
    /// 图片链接
    /// </summary>
    wdLinkTypePicture,
    /// <summary>
    /// 文本文件链接
    /// </summary>
    wdLinkTypeText,
    /// <summary>
    /// 引用链接
    /// </summary>
    wdLinkTypeReference,
    /// <summary>
    /// 包含文件链接
    /// </summary>
    wdLinkTypeInclude,
    /// <summary>
    /// 导入文件链接
    /// </summary>
    wdLinkTypeImport,
    /// <summary>
    /// DDE连接链接
    /// </summary>
    wdLinkTypeDDE,
    /// <summary>
    /// 自动更新的DDE连接链接
    /// </summary>
    wdLinkTypeDDEAuto,
    /// <summary>
    /// 图表链接
    /// </summary>
    wdLinkTypeChart
}