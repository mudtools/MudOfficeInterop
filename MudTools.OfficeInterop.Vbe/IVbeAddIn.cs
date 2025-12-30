//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 表示一个 VBA 加载项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeAddIn : IOfficeObject<IVbeAddIn>, IDisposable
{
    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取或设置加载项的描述文本。
    /// </summary>
    string Description { get; set; }

    /// <summary>
    /// 获取包含所有加载项的 AddIns 集合对象。
    /// </summary>
    IVbeAddins? Collection { get; }

    /// <summary>
    /// 获取加载项的编程标识符（ProgID）。
    /// </summary>
    string ProgId { get; }

    /// <summary>
    /// 获取加载项的全局唯一标识符（GUID）。
    /// </summary>
    string Guid { get; }

    /// <summary>
    /// 获取或设置一个值，指示加载项是否已连接（启用）。
    /// </summary>
    bool Connect { get; set; }

    /// <summary>
    /// 获取或设置与加载项关联的对象（通常为加载项实例）。
    /// </summary>
    object Object { get; set; }
}