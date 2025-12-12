//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定软边效果的类型
/// </summary>
public enum MsoSoftEdgeType
{
    /// <summary>
    /// 混合类型
    /// </summary>
    msoSoftEdgeTypeMixed = -2,
    /// <summary>
    /// 无软边效果
    /// </summary>
    msoSoftEdgeTypeNone = 0,
    /// <summary>
    /// 软边效果类型1
    /// </summary>
    msoSoftEdgeType1 = 1,
    /// <summary>
    /// 软边效果类型2
    /// </summary>
    msoSoftEdgeType2 = 2,
    /// <summary>
    /// 软边效果类型3
    /// </summary>
    msoSoftEdgeType3 = 3,
    /// <summary>
    /// 软边效果类型4
    /// </summary>
    msoSoftEdgeType4 = 4,
    /// <summary>
    /// 软边效果类型5
    /// </summary>
    msoSoftEdgeType5 = 5,
    /// <summary>
    /// 软边效果类型6
    /// </summary>
    msoSoftEdgeType6 = 6
}