//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示同步事件类型的枚举
/// </summary>
public enum MsoSyncEventType
{
    /// <summary>
    /// 下载初始化事件
    /// </summary>
    msoSyncEventDownloadInitiated,

    /// <summary>
    /// 下载成功事件
    /// </summary>
    msoSyncEventDownloadSucceeded,

    /// <summary>
    /// 下载失败事件
    /// </summary>
    msoSyncEventDownloadFailed,

    /// <summary>
    /// 上传初始化事件
    /// </summary>
    msoSyncEventUploadInitiated,

    /// <summary>
    /// 上传成功事件
    /// </summary>
    msoSyncEventUploadSucceeded,

    /// <summary>
    /// 上传失败事件
    /// </summary>
    msoSyncEventUploadFailed,

    /// <summary>
    /// 下载无变化事件
    /// </summary>
    msoSyncEventDownloadNoChange,

    /// <summary>
    /// 离线事件
    /// </summary>
    msoSyncEventOffline
}