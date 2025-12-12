//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本下划线类型
/// </summary>
public enum MsoTextUnderlineType
{

    /// <summary>
    /// 混合下划线类型
    /// </summary>
    msoUnderlineMixed = -2,

    /// <summary>
    /// 无下划线
    /// </summary>
    msoNoUnderline = 0,

    /// <summary>
    /// 仅为单词添加下划线
    /// </summary>
    msoUnderlineWords = 1,

    /// <summary>
    /// 单线下划线
    /// </summary>
    msoUnderlineSingleLine = 2,

    /// <summary>
    /// 双线下划线
    /// </summary>
    msoUnderlineDoubleLine = 3,

    /// <summary>
    /// 粗单线下划线
    /// </summary>
    msoUnderlineHeavyLine = 4,

    /// <summary>
    /// 点状下划线
    /// </summary>
    msoUnderlineDottedLine = 5,

    /// <summary>
    /// 粗点状下划线
    /// </summary>
    msoUnderlineDottedHeavyLine = 6,

    /// <summary>
    /// 虚线下划线
    /// </summary>
    msoUnderlineDashLine = 7,

    /// <summary>
    /// 粗虚线下划线
    /// </summary>
    msoUnderlineDashHeavyLine = 8,

    /// <summary>
    /// 部长虚线下划线
    /// </summary>
    msoUnderlineDashLongLine = 9,

    /// <summary>
    /// 粗长虚线下划线
    /// </summary>
    msoUnderlineDashLongHeavyLine = 10,

    /// <summary>
    /// 点划线下划线
    /// </summary>
    msoUnderlineDotDashLine = 11,

    /// <summary>
    /// 粗点划线下划线
    /// </summary>
    msoUnderlineDotDashHeavyLine = 12,

    /// <summary>
    /// 双点划线下划线
    /// </summary>
    msoUnderlineDotDotDashLine = 13,

    /// <summary>
    /// 粗双点划线下划线
    /// </summary>
    msoUnderlineDotDotDashHeavyLine = 14,

    /// <summary>
    /// 波浪下划线
    /// </summary>
    msoUnderlineWavyLine = 15,

    /// <summary>
    /// 粗波浪下划线
    /// </summary>
    msoUnderlineWavyHeavyLine = 16,

    /// <summary>
    /// 双波浪下划线
    /// </summary>
    msoUnderlineWavyDoubleLine = 17
}