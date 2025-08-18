//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定线条虚线样式的枚举，用于形状或对象的边框线条样式
/// </summary>
public enum MsoLineDashStyle
{
    /// <summary>
    /// 混合虚线样式
    /// </summary>
    msoLineDashStyleMixed = -2,

    /// <summary>
    /// 实线
    /// </summary>
    msoLineSolid = 1,

    /// <summary>
    /// 方块点线
    /// </summary>
    msoLineSquareDot = 2,

    /// <summary>
    /// 圆点线
    /// </summary>
    msoLineRoundDot = 3,

    /// <summary>
    /// 虚线
    /// </summary>
    msoLineDash = 4,

    /// <summary>
    /// 点划线（一点一划）
    /// </summary>
    msoLineDashDot = 5,

    /// <summary>
    /// 双点划线（两点一划）
    /// </summary>
    msoLineDashDotDot = 6,

    /// <summary>
    /// 长虚线
    /// </summary>
    msoLineLongDash = 7,

    /// <summary>
    /// 长虚线点线（一长虚一短点）
    /// </summary>
    msoLineLongDashDot = 8,

    /// <summary>
    /// 长虚线双点线（一长虚两短点）
    /// </summary>
    msoLineLongDashDotDot = 9,

    /// <summary>
    /// 系统定义的虚线
    /// </summary>
    msoLineSysDash = 10,

    /// <summary>
    /// 系统定义的点线
    /// </summary>
    msoLineSysDot = 11,

    /// <summary>
    /// 系统定义的点划线
    /// </summary>
    msoLineSysDashDot = 12
}