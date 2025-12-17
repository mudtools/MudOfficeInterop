//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel 数据透视表筛选类型枚举
/// </summary>
public enum XlPivotFilterType
{
    /// <summary>
    /// 从列表顶部筛选指定数量的值
    /// </summary>
    xlTopCount = 1,

    /// <summary>
    /// 从列表底部筛选指定数量的值
    /// </summary>
    xlBottomCount,

    /// <summary>
    /// 从列表中筛选指定百分比的值
    /// </summary>
    xlTopPercent,

    /// <summary>
    /// 从列表底部筛选指定百分比的值
    /// </summary>
    xlBottomPercent,

    /// <summary>
    /// 列表顶部值的总和
    /// </summary>
    xlTopSum,

    /// <summary>
    /// 列表底部值的总和
    /// </summary>
    xlBottomSum,

    /// <summary>
    /// 筛选所有匹配指定值的值
    /// </summary>
    xlValueEquals,

    /// <summary>
    /// 筛选所有不匹配指定值的值
    /// </summary>
    xlValueDoesNotEqual,

    /// <summary>
    /// 筛选所有大于指定值的值
    /// </summary>
    xlValueIsGreaterThan,

    /// <summary>
    /// 筛选所有大于或等于指定值的值
    /// </summary>
    xlValueIsGreaterThanOrEqualTo,

    /// <summary>
    /// 筛选所有小于指定值的值
    /// </summary>
    xlValueIsLessThan,

    /// <summary>
    /// 筛选所有小于或等于指定值的值
    /// </summary>
    xlValueIsLessThanOrEqualTo,

    /// <summary>
    /// 筛选所有在指定值范围内的值
    /// </summary>
    xlValueIsBetween,

    /// <summary>
    /// 筛选所有不在指定值范围内的值
    /// </summary>
    xlValueIsNotBetween,

    /// <summary>
    /// 筛选所有匹配指定字符串的标题
    /// </summary>
    xlCaptionEquals,

    /// <summary>
    /// 筛选所有不匹配指定字符串的标题
    /// </summary>
    xlCaptionDoesNotEqual,

    /// <summary>
    /// 筛选所有以指定字符串开头的标题
    /// </summary>
    xlCaptionBeginsWith,

    /// <summary>
    /// 筛选所有不以指定字符串开头的标题
    /// </summary>
    xlCaptionDoesNotBeginWith,

    /// <summary>
    /// 筛选所有以指定字符串结尾的标题
    /// </summary>
    xlCaptionEndsWith,

    /// <summary>
    /// 筛选所有不以指定字符串结尾的标题
    /// </summary>
    xlCaptionDoesNotEndWith,

    /// <summary>
    /// 筛选所有包含指定字符串的标题
    /// </summary>
    xlCaptionContains,

    /// <summary>
    /// 筛选所有不包含指定字符串的标题
    /// </summary>
    xlCaptionDoesNotContain,

    /// <summary>
    /// 筛选所有大于指定值的标题
    /// </summary>
    xlCaptionIsGreaterThan,

    /// <summary>
    /// 筛选所有大于或等于指定值的标题
    /// </summary>
    xlCaptionIsGreaterThanOrEqualTo,

    /// <summary>
    /// 筛选所有小于指定值的标题
    /// </summary>
    xlCaptionIsLessThan,

    /// <summary>
    /// 筛选所有小于或等于指定值的标题
    /// </summary>
    xlCaptionIsLessThanOrEqualTo,

    /// <summary>
    /// 筛选所有在指定值范围内的标题
    /// </summary>
    xlCaptionIsBetween,

    /// <summary>
    /// 筛选所有不在指定值范围内的标题
    /// </summary>
    xlCaptionIsNotBetween,

    /// <summary>
    /// 筛选所有匹配指定日期的日期
    /// </summary>
    xlSpecificDate,

    /// <summary>
    /// 筛选所有不匹配指定日期的日期
    /// </summary>
    xlNotSpecificDate,

    /// <summary>
    /// 筛选所有在指定日期之前的日期
    /// </summary>
    xlBefore,

    /// <summary>
    /// 筛选所有在指定日期或之前的日期
    /// </summary>
    xlBeforeOrEqualTo,

    /// <summary>
    /// 筛选所有在指定日期之后的日期
    /// </summary>
    xlAfter,

    /// <summary>
    /// 筛选所有在指定日期或之后的日期
    /// </summary>
    xlAfterOrEqualTo,

    /// <summary>
    /// 筛选所有在指定日期范围内的日期
    /// </summary>
    xlDateBetween,

    /// <summary>
    /// 筛选所有不在指定日期范围内的日期
    /// </summary>
    xlDateNotBetween,

    /// <summary>
    /// 筛选所有明天（次日）的日期
    /// </summary>
    xlDateTomorrow,

    /// <summary>
    /// 筛选所有今天（当前日期）的日期
    /// </summary>
    xlDateToday,

    /// <summary>
    /// 筛选所有昨天（前一日）的日期
    /// </summary>
    xlDateYesterday,

    /// <summary>
    /// 筛选所有下周的日期
    /// </summary>
    xlDateNextWeek,

    /// <summary>
    /// 筛选所有本周的日期
    /// </summary>
    xlDateThisWeek,

    /// <summary>
    /// 筛选所有上周的日期
    /// </summary>
    xlDateLastWeek,

    /// <summary>
    /// 筛选所有下个月的日期
    /// </summary>
    xlDateNextMonth,

    /// <summary>
    /// 筛选所有本月的日期
    /// </summary>
    xlDateThisMonth,

    /// <summary>
    /// 筛选所有上个月的日期
    /// </summary>
    xlDateLastMonth,

    /// <summary>
    /// 筛选所有下个季度的日期
    /// </summary>
    xlDateNextQuarter,

    /// <summary>
    /// 筛选所有本季度的日期
    /// </summary>
    xlDateThisQuarter,

    /// <summary>
    /// 筛选所有上个季度的日期
    /// </summary>
    xlDateLastQuarter,

    /// <summary>
    /// 筛选所有下一年的日期
    /// </summary>
    xlDateNextYear,

    /// <summary>
    /// 筛选所有本年的日期
    /// </summary>
    xlDateThisYear,

    /// <summary>
    /// 筛选所有上一年的日期
    /// </summary>
    xlDateLastYear,

    /// <summary>
    /// 筛选所有在指定日期一年内的值
    /// </summary>
    xlYearToDate,

    /// <summary>
    /// 筛选第一季度内的所有日期
    /// </summary>
    xlAllDatesInPeriodQuarter1,

    /// <summary>
    /// 筛选第二季度内的所有日期
    /// </summary>
    xlAllDatesInPeriodQuarter2,

    /// <summary>
    /// 筛选第三季度内的所有日期
    /// </summary>
    xlAllDatesInPeriodQuarter3,

    /// <summary>
    /// 筛选第四季度内的所有日期
    /// </summary>
    xlAllDatesInPeriodQuarter4,

    /// <summary>
    /// 筛选一月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodJanuary,

    /// <summary>
    /// 筛选二月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodFebruary,

    /// <summary>
    /// 筛选三月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodMarch,

    /// <summary>
    /// 筛选四月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodApril,

    /// <summary>
    /// 筛选五月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodMay,

    /// <summary>
    /// 筛选六月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodJune,

    /// <summary>
    /// 筛选七月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodJuly,

    /// <summary>
    /// 筛选八月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodAugust,

    /// <summary>
    /// 筛选九月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodSeptember,

    /// <summary>
    /// 筛选十月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodOctober,

    /// <summary>
    /// 筛选十一月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodNovember,

    /// <summary>
    /// 筛选十二月份内的所有日期
    /// </summary>
    xlAllDatesInPeriodDecember
}