//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在以黑白模式渲染形状时使用的显示模式
/// </summary>
public enum MsoBlackWhiteMode
{
    /// <summary>
    /// 混合模式（用于属性不一致的情况）
    /// </summary>
    msoBlackWhiteMixed = -2,
    /// <summary>
    /// 自动模式，系统决定如何显示黑白图像
    /// </summary>
    msoBlackWhiteAutomatic = 1,
    /// <summary>
    /// 灰度模式，将彩色图像转换为灰度图像
    /// </summary>
    msoBlackWhiteGrayScale = 2,
    /// <summary>
    /// 浅灰度模式，使用较浅的灰度显示
    /// </summary>
    msoBlackWhiteLightGrayScale = 3,
    /// <summary>
    /// 反转灰度模式，将图像的明暗部分进行反转
    /// </summary>
    msoBlackWhiteInverseGrayScale = 4,
    /// <summary>
    /// 灰色轮廓模式，使用灰色显示轮廓
    /// </summary>
    msoBlackWhiteGrayOutline = 5,
    /// <summary>
    /// 黑色文字和线条模式，文字和线条以黑色显示
    /// </summary>
    msoBlackWhiteBlackTextAndLine = 6,
    /// <summary>
    /// 高对比度模式，使用最大对比度显示图像
    /// </summary>
    msoBlackWhiteHighContrast = 7,
    /// <summary>
    /// 全黑模式，将图像显示为纯黑色
    /// </summary>
    msoBlackWhiteBlack = 8,
    /// <summary>
    /// 全白模式，将图像显示为纯白色
    /// </summary>
    msoBlackWhiteWhite = 9,
    /// <summary>
    /// 不显示模式，不显示该形状
    /// </summary>
    msoBlackWhiteDontShow = 10
}