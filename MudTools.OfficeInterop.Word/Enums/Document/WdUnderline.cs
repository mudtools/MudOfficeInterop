//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定应用于文本的下划线样式
/// </summary>
public enum WdUnderline
{
    /// <summary>
    /// 无下划线
    /// </summary>
    wdUnderlineNone = 0,

    /// <summary>
    /// 单线下划线
    /// </summary>
    wdUnderlineSingle = 1,

    /// <summary>
    /// 仅为单词添加下划线
    /// </summary>
    wdUnderlineWords = 2,

    /// <summary>
    /// 双线下划线
    /// </summary>
    wdUnderlineDouble = 3,

    /// <summary>
    /// 点线下划线
    /// </summary>
    wdUnderlineDotted = 4,

    /// <summary>
    /// 粗下划线
    /// </summary>
    wdUnderlineThick = 6,

    /// <summary>
    /// 虚线下划线
    /// </summary>
    wdUnderlineDash = 7,

    /// <summary>
    /// 点划线下划线
    /// </summary>
    wdUnderlineDotDash = 9,

    /// <summary>
    /// 双点划线下划线
    /// </summary>
    wdUnderlineDotDotDash = 10,

    /// <summary>
    /// 波浪下划线
    /// </summary>
    wdUnderlineWavy = 11,

    /// <summary>
    /// 粗波浪下划线
    /// </summary>
    wdUnderlineWavyHeavy = 27,

    /// <summary>
    /// 粗点线下划线
    /// </summary>
    wdUnderlineDottedHeavy = 20,

    /// <summary>
    /// 粗虚线下划线
    /// </summary>
    wdUnderlineDashHeavy = 23,

    /// <summary>
    /// 粗点划线下划线
    /// </summary>
    wdUnderlineDotDashHeavy = 25,

    /// <summary>
    /// 粗双点划线下划线
    /// </summary>
    wdUnderlineDotDotDashHeavy = 26,

    /// <summary>
    /// 长虚线下划线
    /// </summary>
    wdUnderlineDashLong = 39,

    /// <summary>
    /// 粗长虚线下划线
    /// </summary>
    wdUnderlineDashLongHeavy = 55,

    /// <summary>
    /// 双波浪下划线
    /// </summary>
    wdUnderlineWavyDouble = 43
}