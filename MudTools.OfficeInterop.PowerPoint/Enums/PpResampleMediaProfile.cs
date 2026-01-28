//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定媒体重新采样的配置文件质量。
/// </summary>
public enum PpResampleMediaProfile
{
    /// <summary>
    /// 自定义重新采样配置文件。
    /// </summary>
    ppResampleMediaProfileCustom = 1,

    /// <summary>
    /// 小型重新采样配置文件（中等质量）。
    /// </summary>
    ppResampleMediaProfileSmall,

    /// <summary>
    /// 较小重新采样配置文件（较低质量）。
    /// </summary>
    ppResampleMediaProfileSmaller,

    /// <summary>
    /// 最小重新采样配置文件（最低质量）。
    /// </summary>
    ppResampleMediaProfileSmallest
}