//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 系统信息接口
/// </summary>
public interface IWordSystemInfo
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取操作系统版本
    /// </summary>
    string OSVersion { get; }

    /// <summary>
    /// 获取系统内存大小
    /// </summary>
    long TotalMemory { get; }

    /// <summary>
    /// 获取可用内存大小
    /// </summary>
    long AvailableMemory { get; }

    /// <summary>
    /// 获取处理器数量
    /// </summary>
    int ProcessorCount { get; }

    /// <summary>
    /// 获取系统启动时间
    /// </summary>
    DateTime SystemBootTime { get; }

    /// <summary>
    /// 获取系统运行时间
    /// </summary>
    TimeSpan SystemUptime { get; }
}