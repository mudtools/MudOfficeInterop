//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定项目符号的类型。
/// </summary>
public enum PpBulletType
{
    /// <summary>
    /// 混合项目符号类型。
    /// </summary>
    ppBulletMixed = -2,

    /// <summary>
    /// 无项目符号。
    /// </summary>
    ppBulletNone = 0,

    /// <summary>
    /// 非数字项目符号（如圆点、方块等）。
    /// </summary>
    ppBulletUnnumbered = 1,

    /// <summary>
    /// 数字编号项目符号。
    /// </summary>
    ppBulletNumbered = 2,

    /// <summary>
    /// 图片项目符号。
    /// </summary>
    ppBulletPicture = 3
}