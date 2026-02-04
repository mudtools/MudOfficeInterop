//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中的媒体播放器对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPlayer : IOfficeObject<IPowerPointPlayer, MsPowerPoint.Player>, IDisposable
{
    /// <summary>
    /// 获取创建此播放器的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此播放器的父对象。
    /// </summary>
    /// <value>表示此播放器父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 播放媒体。
    /// </summary>
    void Play();

    /// <summary>
    /// 暂停媒体播放。
    /// </summary>
    void Pause();

    /// <summary>
    /// 停止媒体播放。
    /// </summary>
    void Stop();

    /// <summary>
    /// 转到下一个书签。
    /// </summary>
    void GoToNextBookmark();

    /// <summary>
    /// 转到上一个书签。
    /// </summary>
    void GoToPreviousBookmark();

    /// <summary>
    /// 获取或设置当前播放位置（以毫秒为单位）。
    /// </summary>
    /// <value>表示当前播放位置的整数值。</value>
    int CurrentPosition { get; set; }

    /// <summary>
    /// 获取播放器的当前状态。
    /// </summary>
    /// <value>表示播放器状态的 <see cref="PpPlayerState"/> 枚举值。</value>
    PpPlayerState State { get; }
}