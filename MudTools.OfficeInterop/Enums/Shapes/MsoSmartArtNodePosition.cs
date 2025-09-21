//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定SmartArt节点的相对位置
/// </summary>
public enum MsoSmartArtNodePosition
{
    /// <summary>
    /// 默认位置
    /// </summary>
    msoSmartArtNodeDefault = 1,
    /// <summary>
    /// 当前节点之后
    /// </summary>
    msoSmartArtNodeAfter,
    /// <summary>
    /// 当前节点之前
    /// </summary>
    msoSmartArtNodeBefore,
    /// <summary>
    /// 当前节点上方
    /// </summary>
    msoSmartArtNodeAbove,
    /// <summary>
    /// 当前节点下方
    /// </summary>
    msoSmartArtNodeBelow
}