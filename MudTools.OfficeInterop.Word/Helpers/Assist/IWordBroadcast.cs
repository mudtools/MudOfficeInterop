//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示广播功能，用于管理文档或演示文稿的广播会话。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordBroadcast : IOfficeObject<IWordBroadcast>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取与会者访问广播的 URL。
    /// </summary>
    string AttendeeUrl { get; }

    /// <summary>
    /// 获取广播的当前状态。
    /// </summary>
    MsoBroadcastState State { get; }

    /// <summary>
    /// 获取广播功能的可用能力。
    /// </summary>
    int Capabilities { get; }

    /// <summary>
    /// 获取主持人服务的 URL。
    /// </summary>
    string PresenterServiceUrl { get; }

    /// <summary>
    /// 获取广播会话的唯一标识符。
    /// </summary>
    string SessionID { get; }

    /// <summary>
    /// 开始新的广播会话。
    /// </summary>
    /// <param name="serverUrl">广播服务器的 URL。</param>
    void Start(string serverUrl);

    /// <summary>
    /// 暂停当前广播会话。
    /// </summary>
    void Pause();

    /// <summary>
    /// 恢复暂停的广播会话。
    /// </summary>
    void Resume();

    /// <summary>
    /// 结束当前广播会话。
    /// </summary>
    void End();

    /// <summary>
    /// 向广播会话添加会议笔记。
    /// </summary>
    /// <param name="notesUrl">会议笔记的 URL。</param>
    /// <param name="notesWacUrl">会议笔记 WAC（Web App Companion）的 URL。</param>
    void AddMeetingNotes(string notesUrl, string notesWacUrl);
}