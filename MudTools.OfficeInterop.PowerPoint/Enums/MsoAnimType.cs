//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定动画的类型。
/// </summary>
public enum MsoAnimType
{
    /// <summary>
    /// 混合动画类型。
    /// </summary>
    msoAnimTypeMixed = -2,

    /// <summary>
    /// 无动画类型。
    /// </summary>
    msoAnimTypeNone = 0,

    /// <summary>
    /// 运动动画类型。
    /// </summary>
    msoAnimTypeMotion = 1,

    /// <summary>
    /// 颜色动画类型。
    /// </summary>
    msoAnimTypeColor = 2,

    /// <summary>
    /// 缩放动画类型。
    /// </summary>
    msoAnimTypeScale = 3,

    /// <summary>
    /// 旋转动画类型。
    /// </summary>
    msoAnimTypeRotation = 4,

    /// <summary>
    /// 属性动画类型。
    /// </summary>
    msoAnimTypeProperty = 5,

    /// <summary>
    /// 命令动画类型。
    /// </summary>
    msoAnimTypeCommand = 6,

    /// <summary>
    /// 滤镜动画类型。
    /// </summary>
    msoAnimTypeFilter = 7,

    /// <summary>
    /// 设置动画类型。
    /// </summary>
    msoAnimTypeSet = 8
}