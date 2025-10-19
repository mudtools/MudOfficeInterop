namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在文档中插入的分节符、分页符或分栏符的类型
/// </summary>
public enum WdBreakType
{

    /// <summary>
    /// 下一页分节符 - 在下一页开始新节
    /// </summary>
    wdSectionBreakNextPage = 2,

    /// <summary>
    /// 连续分节符 - 在同一页面开始新节
    /// </summary>
    wdSectionBreakContinuous,

    /// <summary>
    /// 偶数页分节符 - 在下一个偶数页开始新节
    /// </summary>
    wdSectionBreakEvenPage,

    /// <summary>
    /// 奇数页分节符 - 在下一个奇数页开始新节
    /// </summary>
    wdSectionBreakOddPage,

    /// <summary>
    /// 行结束符 - 结束当前行并强制文本移至下一行
    /// </summary>
    wdLineBreak,

    /// <summary>
    /// 分页符 - 在插入点处分页
    /// </summary>
    wdPageBreak,

    /// <summary>
    /// 分栏符 - 在插入点处分栏
    /// </summary>
    wdColumnBreak,

    /// <summary>
    /// 左侧清除行结束符 - 清除左侧的文本环绕
    /// </summary>
    wdLineBreakClearLeft,

    /// <summary>
    /// 右侧清除行结束符 - 清除右侧的文本环绕
    /// </summary>
    wdLineBreakClearRight,

    /// <summary>
    /// 文本环绕分隔符 - 结束当前行并将文本继续放在图片、表格或其他项目的下方
    /// </summary>
    wdTextWrappingBreak
}