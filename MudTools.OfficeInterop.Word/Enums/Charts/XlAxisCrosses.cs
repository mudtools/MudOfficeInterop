//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定另一个轴在指定轴上的交叉点
/// </summary>
public enum XlAxisCrosses
{
    /// <summary>
    /// 由 Microsoft Word 自动设置轴交叉点
    /// </summary>
    xlAxisCrossesAutomatic = -4105,

    /// <summary>
    /// 通过 CrossesAt 属性指定轴交叉点
    /// </summary>
    xlAxisCrossesCustom = -4114,

    /// <summary>
    /// 轴在最大值处交叉
    /// </summary>
    xlAxisCrossesMaximum = 2,

    /// <summary>
    /// 轴在最小值处交叉
    /// </summary>
    xlAxisCrossesMinimum = 4
}