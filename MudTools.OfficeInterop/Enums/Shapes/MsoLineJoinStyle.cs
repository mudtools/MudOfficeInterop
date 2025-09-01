//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定线条连接样式的枚举类型，用于定义形状路径中线段相交处的连接方式
/// </summary>
public enum MsoLineJoinStyle
{
    /// <summary>
    /// 混合线条连接样式（通常用于表示未设置或混合了多种样式）
    /// </summary>
    msoLineJoinMixed = -2,

    /// <summary>
    /// 圆角连接，线段相交处以圆弧过渡
    /// </summary>
    msoLineJoinRound = 1,

    /// <summary>
    /// 斜角连接，线段相交处以斜切方式连接
    /// </summary>
    msoLineJoinBevel = 2,

    /// <summary>
    /// 直角连接，线段相交处以直角方式连接
    /// </summary>
    msoLineJoinMiter = 3
}