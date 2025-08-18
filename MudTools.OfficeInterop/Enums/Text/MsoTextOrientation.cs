//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本方向的枚举，用于控制文本的排列方向
/// </summary>
public enum MsoTextOrientation
{
    /// <summary>
    /// 混合文本方向
    /// </summary>
    msoTextOrientationMixed = -2,

    /// <summary>
    /// 水平方向
    /// </summary>
    msoTextOrientationHorizontal = 1,

    /// <summary>
    /// 向上方向（文字从下往上）
    /// </summary>
    msoTextOrientationUpward = 2,

    /// <summary>
    /// 向下方向（文字从上往下）
    /// </summary>
    msoTextOrientationDownward = 3,

    /// <summary>
    /// 远东垂直方向（垂直排列，适合中文等语言）
    /// </summary>
    msoTextOrientationVerticalFarEast = 4,

    /// <summary>
    /// 垂直方向（从右到左垂直排列）
    /// </summary>
    msoTextOrientationVertical = 5,

    /// <summary>
    /// 远东水平旋转方向（水平排列但字符旋转90度）
    /// </summary>
    msoTextOrientationHorizontalRotatedFarEast = 6
}