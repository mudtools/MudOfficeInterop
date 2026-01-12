namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 模板类型枚举
/// 用于指定文档所使用的模板类型
/// </summary>
public enum WdTemplateType
{
    /// <summary>
    /// 正常模板 - Word 的默认模板
    /// </summary>
    wdNormalTemplate,
    
    /// <summary>
    /// 全局模板 - 在所有文档中可用的全局模板
    /// </summary>
    wdGlobalTemplate,
    
    /// <summary>
    /// 附加模板 - 附加到特定文档的模板
    /// </summary>
    wdAttachedTemplate
}