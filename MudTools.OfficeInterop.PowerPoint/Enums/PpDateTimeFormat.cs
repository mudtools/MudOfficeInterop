//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定日期和时间的显示格式。
/// </summary>
public enum PpDateTimeFormat
{
    /// <summary>
    /// 混合日期时间格式。
    /// </summary>
    ppDateTimeFormatMixed = -2,

    /// <summary>
    /// 月/日/年（如 3/14/2024）。
    /// </summary>
    ppDateTimeMdyy = 1,

    /// <summary>
    /// 星期几 月份 日 年（如 Tuesday March 14 2024）。
    /// </summary>
    ppDateTimeddddMMMMddyyyy = 2,

    /// <summary>
    /// 日 月份 年（如 14 March 2024）。
    /// </summary>
    ppDateTimedMMMMyyyy = 3,

    /// <summary>
    /// 月份 日 年（如 March 14 2024）。
    /// </summary>
    ppDateTimeMMMMdyyyy = 4,

    /// <summary>
    /// 日 月简称 年（如 14 Mar 24）。
    /// </summary>
    ppDateTimedMMMyy = 5,

    /// <summary>
    /// 月份 年（如 March 2024）。
    /// </summary>
    ppDateTimeMMMMyy = 6,

    /// <summary>
    /// 月/年（如 03/24）。
    /// </summary>
    ppDateTimeMMyy = 7,

    /// <summary>
    /// 月/日/年 时:分（如 03/14/24 14:30）。
    /// </summary>
    ppDateTimeMMddyyHmm = 8,

    /// <summary>
    /// 月/日/年 时:分 AM/PM（如 03/14/24 2:30 PM）。
    /// </summary>
    ppDateTimeMMddyyhmmAMPM = 9,

    /// <summary>
    /// 时:分（24 小时制，如 14:30）。
    /// </summary>
    ppDateTimeHmm = 10,

    /// <summary>
    /// 时:分:秒（24 小时制，如 14:30:45）。
    /// </summary>
    ppDateTimeHmmss = 11,

    /// <summary>
    /// 时:分 AM/PM（如 2:30 PM）。
    /// </summary>
    ppDateTimehmmAMPM = 12,

    /// <summary>
    /// 时:分:秒 AM/PM（如 2:30:45 PM）。
    /// </summary>
    ppDateTimehmmssAMPM = 13,

    /// <summary>
    /// 自动推断格式。
    /// </summary>
    ppDateTimeFigureOut = 14,

    /// <summary>
    /// 用户区域设置格式 1。
    /// </summary>
    ppDateTimeUAQ1 = 15,

    /// <summary>
    /// 用户区域设置格式 2。
    /// </summary>
    ppDateTimeUAQ2 = 16,

    /// <summary>
    /// 用户区域设置格式 3。
    /// </summary>
    ppDateTimeUAQ3 = 17,

    /// <summary>
    /// 用户区域设置格式 4。
    /// </summary>
    ppDateTimeUAQ4 = 18,

    /// <summary>
    /// 用户区域设置格式 5。
    /// </summary>
    ppDateTimeUAQ5 = 19,

    /// <summary>
    /// 用户区域设置格式 6。
    /// </summary>
    ppDateTimeUAQ6 = 20,

    /// <summary>
    /// 用户区域设置格式 7。
    /// </summary>
    ppDateTimeUAQ7 = 21
}