//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示提供有关 Windows 注册表中注册的 COM 加载项信息的 COMAddIn 对象集合。
/// 此接口提供对 COM 加载项集合的管理功能，包括检索、更新和枚举 COM 加载项。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeCOMAddIns : IEnumerable<IOfficeCOMAddIn?>, IOfficeObject<IOfficeCOMAddIns, MsCore.COMAddIns>, IDisposable
{
    /// <summary>
    /// 返回指定 COMAddIns 集合中的一个成员。
    /// </summary>
    /// <param name="index">必需的对象。可以是返回 COMAddIns 集合中该位置的 COM 加载项的序数值，
    /// 也可以是表示指定 COM 加载项的 ProgID 的字符串值。</param>
    /// <returns>指定索引处的 COM 加载项对象。</returns>
    IOfficeCOMAddIn? this[int index] { get; }

    /// <summary>
    /// 返回指定 COMAddIns 集合中的一个成员。
    /// </summary>
    /// <param name="progID">必需的对象。可以是返回 COMAddIns 集合中该位置的 COM 加载项的序数值，
    /// 也可以是表示指定 COM 加载项的 ProgID 的字符串值。</param>
    /// <returns>指定索引处的 COM 加载项对象。</returns>
    IOfficeCOMAddIn? this[string progID] { get; }

    /// <summary>
    /// 获取指定集合中的项数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据 Windows 注册表中存储的加载项列表更新 COMAddIns 集合的内容。
    /// </summary>
    void Update();

    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 内部使用。
    /// </summary>
    /// <param name="modal">模态状态。</param>
    void SetAppModal(bool modal);
}