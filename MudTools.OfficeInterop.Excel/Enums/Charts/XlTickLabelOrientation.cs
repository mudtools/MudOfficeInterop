//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 刻度标签方向枚举
/// 用于指定图表坐标轴上刻度标签的文字方向
/// </summary>
public enum XlTickLabelOrientation
{
    /// <summary>
    /// 自动方向
    /// Excel根据图表类型自动选择合适的标签方向
    /// </summary>
    xlTickLabelOrientationAutomatic = -4105,

    /// <summary>
    /// 向下倾斜
    /// 标签文字从左上到右下倾斜排列
    /// </summary>
    xlTickLabelOrientationDownward = -4170,

    /// <summary>
    /// 水平方向
    /// 标签文字水平排列
    /// </summary>
    xlTickLabelOrientationHorizontal = -4128,

    /// <summary>
    /// 向上倾斜
    /// 标签文字从左下到右上倾斜排列
    /// </summary>
    xlTickLabelOrientationUpward = -4171,

    /// <summary>
    /// 垂直方向
    /// 标签文字垂直排列
    /// </summary>
    xlTickLabelOrientationVertical = -4166
}
