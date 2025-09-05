namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定首选宽度类型，用于确定对象的宽度计算方式
/// </summary>
public enum WdPreferredWidthType
{
    /// <summary>
    /// 自动宽度 - 系统根据内容自动调整宽度
    /// </summary>
    wdPreferredWidthAuto = 1,

    /// <summary>
    /// 百分比宽度 - 按照父容器宽度的百分比设置宽度
    /// </summary>
    wdPreferredWidthPercent,

    /// <summary>
    /// 点宽度 - 按照点（point）单位设置具体宽度值
    /// </summary>
    wdPreferredWidthPoints
}