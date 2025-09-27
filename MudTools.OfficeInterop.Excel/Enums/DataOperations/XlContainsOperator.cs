namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定在Excel数据操作中用于文本匹配的运算符类型
/// </summary>
public enum XlContainsOperator
{
    /// <summary>
    /// 表示包含指定文本
    /// </summary>
    xlContains,
    /// <summary>
    /// 表示不包含指定文本
    /// </summary>
    xlDoesNotContain,
    /// <summary>
    /// 表示以指定文本开头
    /// </summary>
    xlBeginsWith,
    /// <summary>
    /// 表示以指定文本结尾
    /// </summary>
    xlEndsWith
}