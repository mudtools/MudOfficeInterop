//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定基于源范围的内容，目标范围应如何填充
/// </summary>
public enum XlAutoFillType
{
    /// <summary>
    /// 将值和格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillCopy = 1,

    /// <summary>
    /// 将源范围中的星期几名称扩展到目标范围。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillDays = 5,

    /// <summary>
    /// Excel确定用于填充目标范围的值和格式
    /// </summary>
    xlFillDefault = 0,

    /// <summary>
    /// 仅将格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillFormats = 3,

    /// <summary>
    /// 将源范围中的月份名称扩展到目标范围。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillMonths = 7,

    /// <summary>
    /// 将源范围中的值作为序列扩展到目标范围（例如，'1, 2'扩展为'3, 4, 5'）。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillSeries = 2,

    /// <summary>
    /// 仅将值从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillValues = 4,

    /// <summary>
    /// 将源范围中的工作日名称扩展到目标范围。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillWeekdays = 6,

    /// <summary>
    /// 将源范围中的年份扩展到目标范围。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlFillYears = 8,

    /// <summary>
    /// 将源范围中的数值扩展到目标范围，假设源范围内数字之间的关系是乘法的（例如，'1, 2,'扩展为'4, 8, 16'，假设每个数字都是前一个数字乘以某个值的结果）。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlGrowthTrend = 10,

    /// <summary>
    /// 将源范围中的数值扩展到目标范围，假设数字之间的关系是加法的（例如，'1, 2,'扩展为'3, 4, 5'，假设每个数字都是前一个数字加上某个值的结果）。格式从源范围复制到目标范围，必要时重复
    /// </summary>
    xlLinearTrend = 9,

    /// <summary>
    /// 快速填充
    /// </summary>
    xlFlashFill = 11
}