namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定公式中使用的标签类型
/// </summary>
public enum XlFormulaLabel
{
    /// <summary>
    /// 不使用标签
    /// </summary>
    xlNoLabels = -4142,

    /// <summary>
    /// 使用行标签
    /// </summary>
    xlRowLabels = 1,

    /// <summary>
    /// 使用列标签
    /// </summary>
    xlColumnLabels = 2,

    /// <summary>
    /// 同时使用行标签和列标签
    /// </summary>
    xlMixedLabels = 3
}