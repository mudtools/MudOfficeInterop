//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// VBE VBComponents 集合对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBComponents 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsVb"), ItemIndex]
public interface IVbeVBComponents : IEnumerable<IVbeVBComponent?>, IOfficeObject<IVbeVBComponents>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 VB 组件集合中的组件数量
    /// 对应 VBComponents.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的 VB 组件对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">组件索引（从1开始）</param>
    /// <returns>VB 组件对象</returns>
    IVbeVBComponent? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的 VB 组件对象
    /// </summary>
    /// <param name="name">组件名称</param>
    /// <returns>VB 组件对象</returns>
    IVbeVBComponent? this[string name] { get; }

    /// <summary>
    /// 获取 VB 组件集合所在的父对象（通常是 VBProject）
    /// 对应 VBComponents.Parent 属性
    /// </summary>
    object? Parent { get; }

    #endregion

    /// <summary>
    /// 从 VB 组件集合中移除指定的组件
    /// </summary>
    /// <param name="vbComponent">要移除的 VB 组件对象</param>
    void Remove(IVbeVBComponent vbComponent);

    /// <summary>
    /// 添加指定类型的 VB 组件到集合中
    /// </summary>
    /// <param name="componentType">组件类型，可选值包括：vbext_ct_StdModule(标准模块)、vbext_ct_ClassModule(类模块)、vbext_ct_MSForm(MS表单)、vbext_ct_ActiveXDesigner(ActiveX设计器)、vbext_ct_Document(文档)</param>
    /// <returns>新添加的 VB 组件对象，如果添加失败则返回 null</returns>
    IVbeVBComponent? Add(vbext_ComponentType componentType);

    /// <summary>
    /// 导入指定文件到 VB 组件集合中
    /// </summary>
    /// <param name="fileName">要导入的文件路径</param>
    /// <returns>导入的 VB 组件对象，如果导入失败则返回 null</returns>
    IVbeVBComponent? Import(string fileName);

    /// <summary>
    /// 添加自定义组件
    /// </summary>
    /// <param name="progId">组件的 ProgID</param>
    /// <returns>新添加的 VB 组件对象，如果添加失败则返回 null</returns>
    IVbeVBComponent? AddCustom(string progId);

    /// <summary>
    /// 添加 MS Forms 设计器组件
    /// </summary>
    /// <param name="index">组件索引</param>
    /// <returns>新添加的 VB 组件对象，如果添加失败则返回 null</returns>
    IVbeVBComponent? AddMTDesigner(int index = 0);
}