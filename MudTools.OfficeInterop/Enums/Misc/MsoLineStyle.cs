//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定线条样式的枚举，用于形状或对象的边框线条类型
/// </summary>
public enum MsoLineStyle
{
    /// <summary>
    /// 混合线条样式
    /// </summary>
    msoLineStyleMixed = -2,

    /// <summary>
    /// 单线
    /// </summary>
    msoLineSingle = 1,

    /// <summary>
    /// 细细线（非常细的双线）
    /// </summary>
    msoLineThinThin = 2,

    /// <summary>
    /// 细粗线（细线在前，粗线在后）
    /// </summary>
    msoLineThinThick = 3,

    /// <summary>
    /// 粗细线（粗线在前，细线在后）
    /// </summary>
    msoLineThickThin = 4,

    /// <summary>
    /// 粗细细线（粗线在中间，两侧为细线）
    /// </summary>
    msoLineThickBetweenThin = 5
}