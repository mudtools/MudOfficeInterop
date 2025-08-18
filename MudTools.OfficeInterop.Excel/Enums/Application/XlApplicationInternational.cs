//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 国际化设置枚举
/// 用于获取Excel应用程序的国际化设置信息，如区域设置、日期格式、货币符号等
/// </summary>
public enum XlApplicationInternational
{
    /// <summary>
    /// 24小时制时钟
    /// 指示系统是否使用24小时制时钟格式
    /// </summary>
    xl24HourClock = 33,
    
    /// <summary>
    /// 4位数年份
    /// 指示日期格式中是否使用4位数表示年份
    /// </summary>
    xl4DigitYears = 43,
    
    /// <summary>
    /// 替代数组分隔符
    /// 用于分隔数组元素的替代字符
    /// </summary>
    xlAlternateArraySeparator = 16,
    
    /// <summary>
    /// 列分隔符
    /// 用于分隔列的字符
    /// </summary>
    xlColumnSeparator = 14,
    
    /// <summary>
    /// 国家代码
    /// 系统当前设置的国家或地区代码
    /// </summary>
    xlCountryCode = 1,
    
    /// <summary>
    /// 国家设置
    /// 系统当前的国家或地区设置
    /// </summary>
    xlCountrySetting = 2,
    
    /// <summary>
    /// 货币前缀
    /// 指示货币符号是否显示在数值之前
    /// </summary>
    xlCurrencyBefore = 37,
    
    /// <summary>
    /// 货币代码
    /// 当前设置的货币符号代码
    /// </summary>
    xlCurrencyCode = 25,
    
    /// <summary>
    /// 货币小数位数
    /// 货币数值中小数点后的位数
    /// </summary>
    xlCurrencyDigits = 27,
    
    /// <summary>
    /// 货币前导零
    /// 指示货币数值是否显示前导零
    /// </summary>
    xlCurrencyLeadingZeros = 40,
    
    /// <summary>
    /// 货币负号
    /// 表示负货币数值的符号
    /// </summary>
    xlCurrencyMinusSign = 38,
    
    /// <summary>
    /// 货币负数格式
    /// 负数货币数值的显示格式
    /// </summary>
    xlCurrencyNegative = 28,
    
    /// <summary>
    /// 货币前空格
    /// 指示货币符号与数值之间是否包含空格
    /// </summary>
    xlCurrencySpaceBefore = 36,
    
    /// <summary>
    /// 货币尾随零
    /// 指示货币数值的小数部分是否显示尾随零
    /// </summary>
    xlCurrencyTrailingZeros = 39,
    
    /// <summary>
    /// 日期顺序
    /// 系统中日期的显示顺序（如月/日/年或日/月/年）
    /// </summary>
    xlDateOrder = 32,
    
    /// <summary>
    /// 日期分隔符
    /// 用于分隔日期各部分的字符（如"/"或"-"）
    /// </summary>
    xlDateSeparator = 17,
    
    /// <summary>
    /// 天代码
    /// 表示日期中"天"部分的代码
    /// </summary>
    xlDayCode = 21,
    
    /// <summary>
    /// 天前导零
    /// 指示日期中的天数是否以两位数显示（不足两位时前面补零）
    /// </summary>
    xlDayLeadingZero = 42,
    
    /// <summary>
    /// 小数分隔符
    /// 用于分隔整数和小数部分的字符（如"."或","）
    /// </summary>
    xlDecimalSeparator = 3,
    
    /// <summary>
    /// 常规格式名称
    /// 系统默认的常规数字格式名称
    /// </summary>
    xlGeneralFormatName = 26,
    
    /// <summary>
    /// 小时代码
    /// 表示时间中"小时"部分的代码
    /// </summary>
    xlHourCode = 22,
    
    /// <summary>
    /// 左花括号
    /// 系统使用的左花括号字符"{"
    /// </summary>
    xlLeftBrace = 12,
    
    /// <summary>
    /// 左方括号
    /// 系统使用的左方括号字符"["
    /// </summary>
    xlLeftBracket = 10,
    
    /// <summary>
    /// 列表分隔符
    /// 用于分隔列表项的字符（如逗号或分号）
    /// </summary>
    xlListSeparator = 5,
    
    /// <summary>
    /// 小写列字母
    /// 表示列标题的小写字母形式
    /// </summary>
    xlLowerCaseColumnLetter = 9,
    
    /// <summary>
    /// 小写行字母
    /// 表示行标题的小写字母形式
    /// </summary>
    xlLowerCaseRowLetter = 8,
    
    /// <summary>
    /// 月日年顺序
    /// 指示日期是否按照月-日-年的顺序显示
    /// </summary>
    xlMDY = 44,
    
    /// <summary>
    /// 公制单位
    /// 指示系统是否使用公制单位
    /// </summary>
    xlMetric = 35,
    
    /// <summary>
    /// 分钟代码
    /// 表示时间中"分钟"部分的代码
    /// </summary>
    xlMinuteCode = 23,
    
    /// <summary>
    /// 月代码
    /// 表示日期中"月"部分的代码
    /// </summary>
    xlMonthCode = 20,
    
    /// <summary>
    /// 月前导零
    /// 指示日期中的月份是否以两位数显示（不足两位时前面补零）
    /// </summary>
    xlMonthLeadingZero = 41,
    
    /// <summary>
    /// 月份名称字符数
    /// 月份名称显示的字符数
    /// </summary>
    xlMonthNameChars = 30,
    
    /// <summary>
    /// 非货币数字位数
    /// 非货币数值中小数点后的位数
    /// </summary>
    xlNoncurrencyDigits = 29,
    
    /// <summary>
    /// 非英语函数
    /// 指示是否使用本地语言的函数名称而非英语函数名称
    /// </summary>
    xlNonEnglishFunctions = 34,
    
    /// <summary>
    /// 右花括号
    /// 系统使用的右花括号字符"}"
    /// </summary>
    xlRightBrace = 13,
    
    /// <summary>
    /// 右方括号
    /// 系统使用的右方括号字符"]"
    /// </summary>
    xlRightBracket = 11,
    
    /// <summary>
    /// 行分隔符
    /// 用于分隔行的字符
    /// </summary>
    xlRowSeparator = 15,
    
    /// <summary>
    /// 秒代码
    /// 表示时间中"秒"部分的代码
    /// </summary>
    xlSecondCode = 24,
    
    /// <summary>
    /// 千位分隔符
    /// 用于分隔千位的字符（如","或"."）
    /// </summary>
    xlThousandsSeparator = 4,
    
    /// <summary>
    /// 时间前导零
    /// 指示时间数值是否显示前导零
    /// </summary>
    xlTimeLeadingZero = 45,
    
    /// <summary>
    /// 时间分隔符
    /// 用于分隔时间各部分的字符（如":"）
    /// </summary>
    xlTimeSeparator = 18,
    
    /// <summary>
    /// 大写列字母
    /// 表示列标题的大写字母形式
    /// </summary>
    xlUpperCaseColumnLetter = 7,
    
    /// <summary>
    /// 大写行字母
    /// 表示行标题的大写字母形式
    /// </summary>
    xlUpperCaseRowLetter = 6,
    
    /// <summary>
    /// 星期名称字符数
    /// 星期名称显示的字符数
    /// </summary>
    xlWeekdayNameChars = 31,
    
    /// <summary>
    /// 年代码
    /// 表示日期中"年"部分的代码
    /// </summary>
    xlYearCode = 19
}