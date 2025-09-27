namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中可用的时间段选项，用于数据操作和筛选
/// </summary>
public enum XlTimePeriods
{
    /// <summary>
    /// 今天
    /// </summary>
    xlToday,
    /// <summary>
    /// 昨天
    /// </summary>
    xlYesterday,
    /// <summary>
    /// 最近7天
    /// </summary>
    xlLast7Days,
    /// <summary>
    /// 本周
    /// </summary>
    xlThisWeek,
    /// <summary>
    /// 上周
    /// </summary>
    xlLastWeek,
    /// <summary>
    /// 上个月
    /// </summary>
    xlLastMonth,
    /// <summary>
    /// 明天
    /// </summary>
    xlTomorrow,
    /// <summary>
    /// 下周
    /// </summary>
    xlNextWeek,
    /// <summary>
    /// 下个月
    /// </summary>
    xlNextMonth,
    /// <summary>
    /// 本月
    /// </summary>
    xlThisMonth
}