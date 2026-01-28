//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 播放状态枚举
/// </summary>
public enum PpPlayState
{
    /// <summary>
    /// 已停止
    /// </summary>
    ppPlayStateStopped = 0,

    /// <summary>
    /// 已暂停
    /// </summary>
    ppPlayStatePaused = 1,

    /// <summary>
    /// 正在播放
    /// </summary>
    ppPlayStatePlaying = 2,

    /// <summary>
    /// 扫描中
    /// </summary>
    ppPlayStateScanning = 3,

    /// <summary>
    /// 连接中
    /// </summary>
    ppPlayStateConnecting = 4,

    /// <summary>
    /// 缓冲中
    /// </summary>
    ppPlayStateBuffering = 5,

    /// <summary>
    /// 等待中
    /// </summary>
    ppPlayStateWaiting = 6,

    /// <summary>
    /// 媒体已结束
    /// </summary>
    ppPlayStateMediaEnded = 7,

    /// <summary>
    /// 转换中
    /// </summary>
    ppPlayStateTransitioning = 8
}