//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本基线对齐方式
/// </summary>
public enum MsoBaselineAlignment
{
    /// <summary>
    /// 混合基线对齐模式
    /// </summary>
    msoBaselineAlignMixed = -2,

    /// <summary>
    /// 基线对齐
    /// </summary>
    msoBaselineAlignBaseline = 1,

    /// <summary>
    /// 顶部对齐
    /// </summary>
    msoBaselineAlignTop = 2,

    /// <summary>
    /// 居中对齐
    /// </summary>
    msoBaselineAlignCenter = 3,

    /// <summary>
    /// 远东语言50%对齐（专门针对东亚文字的特殊对齐方式）
    /// </summary>
    msoBaselineAlignFarEast50 = 4,

    /// <summary>
    /// 自动对齐
    /// </summary>
    msoBaselineAlignAuto = 5
}