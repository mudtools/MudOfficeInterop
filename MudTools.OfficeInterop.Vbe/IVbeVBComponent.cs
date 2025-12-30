//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE VBComponent 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBComponent 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeVBComponent : IOfficeObject<IVbeVBComponent>, IDisposable
{
    /// <summary>
    /// 获取一个值，指示组件是否已保存。
    /// </summary>
    bool Saved { get; }

    /// <summary>
    /// 获取或设置组件的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取组件的设计器对象。
    /// </summary>
    object Designer { get; }

    /// <summary>
    /// 获取组件的代码模块。
    /// </summary>
    IVbeCodeModule? CodeModule { get; }

    /// <summary>
    /// 获取组件的类型（如标准模块、类模块、窗体等）。
    /// </summary>
    vbext_ComponentType Type { get; }

    /// <summary>
    /// 将组件导出到指定文件。
    /// </summary>
    /// <param name="fileName">导出文件的路径和名称。</param>
    void Export(string fileName);

    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取包含此组件的组件集合。
    /// </summary>
    IVbeVBComponents? Collection { get; }

    /// <summary>
    /// 获取一个值，指示组件是否有打开的设计器窗口。
    /// </summary>
    bool HasOpenDesigner { get; }

    /// <summary>
    /// 获取组件设计器窗口。
    /// </summary>
    /// <returns>设计器窗口对象。</returns>
    IVbeWindow? DesignerWindow();

    /// <summary>
    /// 激活组件（使其成为活动组件）。
    /// </summary>
    void Activate();

    /// <summary>
    /// 获取组件设计器的标识符。
    /// </summary>
    string DesignerID { get; }

}
