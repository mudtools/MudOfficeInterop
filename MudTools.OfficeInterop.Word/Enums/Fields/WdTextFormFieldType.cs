namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定文本表单域的类型
/// </summary>
public enum WdTextFormFieldType
{

    /// <summary>
    /// 常规文本
    /// </summary>
    wdRegularText,

    /// <summary>
    /// 数字文本
    /// </summary>
    wdNumberText,

    /// <summary>
    /// 日期文本
    /// </summary>
    wdDateText,

    /// <summary>
    /// 当前日期文本
    /// </summary>
    wdCurrentDateText,

    /// <summary>
    /// 当前时间文本
    /// </summary>
    wdCurrentTimeText,

    /// <summary>
    /// 计算文本
    /// </summary>
    wdCalculationText
}