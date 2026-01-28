//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定配色方案中的颜色索引。
/// </summary>
public enum PpColorSchemeIndex
{
    /// <summary>
    /// 混合配色方案颜色。
    /// </summary>
    ppSchemeColorMixed = -2,

    /// <summary>
    /// 非配色方案颜色。
    /// </summary>
    ppNotSchemeColor = 0,

    /// <summary>
    /// 背景颜色。
    /// </summary>
    ppBackground = 1,

    /// <summary>
    /// 前景颜色。
    /// </summary>
    ppForeground = 2,

    /// <summary>
    /// 阴影颜色。
    /// </summary>
    ppShadow = 3,

    /// <summary>
    /// 标题颜色。
    /// </summary>
    ppTitle = 4,

    /// <summary>
    /// 填充颜色。
    /// </summary>
    ppFill = 5,

    /// <summary>
    /// 强调颜色1。
    /// </summary>
    ppAccent1 = 6,

    /// <summary>
    /// 强调颜色2。
    /// </summary>
    ppAccent2 = 7,

    /// <summary>
    /// 强调颜色3。
    /// </summary>
    ppAccent3 = 8
}