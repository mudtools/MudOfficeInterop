//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE VBProject 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBProject 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeVBProject : IOfficeObject<IVbeVBProject>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 VB 项目的名称
    /// 对应 VBProject.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取 VB 项目的类型
    /// 对应 VBProject.Type 属性
    /// </summary>
    vbext_ProjectType Type { get; }

    /// <summary>
    /// 获取 VB 项目的父对象（通常是 VBProjects 集合）
    /// 对应 VBProject.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取 VB 项目的完整路径（如果已保存）
    /// 对应 VBProject.FileName 属性
    /// </summary>
    string FileName { get; }

    /// <summary>
    /// 获取 VB 项目的构建文件名
    /// 对应 VBProject.BuildFileName 属性
    /// </summary>
    string BuildFileName { get; set; }

    /// <summary>
    /// 获取 VB 项目的 VBE 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取 VB 项目的集合对象
    /// 对应 VBProject.Collection 属性
    /// </summary>
    IVbeVBProjects? Collection { get; }

    /// <summary>
    /// 获取 VB 项目的描述
    /// 对应 VBProject.Description 属性
    /// </summary>
    string Description { get; set; }

    /// <summary>
    /// 获取 VB 项目的帮助文件路径
    /// 对应 VBProject.HelpFile 属性
    /// </summary>
    string HelpFile { get; set; }

    /// <summary>
    /// 获取 VB 项目的保存状态
    /// 对应 VBProject.Saved 属性
    /// </summary>
    bool Saved { get; }

    /// <summary>
    /// 获取 VB 项目的帮助上下文 ID
    /// 对应 VBProject.HelpContextID 属性
    /// </summary>
    int HelpContextID { get; set; }

    /// <summary>
    /// 获取 VB 项目的模式（设计模式、运行模式等）
    /// 对应 VBProject.Mode 属性
    /// </summary>
    vbext_VBAMode Mode { get; }

    /// <summary>
    /// 获取 VB 项目的保护状态
    /// 对应 VBProject.Protection 属性
    /// </summary>
    vbext_ProjectProtection Protection { get; }
    #endregion

    #region 核心对象集合
    /// <summary>
    /// 获取 VB 项目的组件集合
    /// 对应 VBProject.VBComponents 属性
    /// </summary>
    IVbeVBComponents? VBComponents { get; }

    /// <summary>
    /// 获取 VB 项目的引用集合
    /// 对应 VBProject.References 属性
    /// </summary>
    IVbeReferences? References { get; }
    #endregion

    #region 操作方法 
    /// <summary>
    /// 另存 VB 项目为
    /// 对应 VBProject.SaveAs 方法
    /// </summary>
    /// <param name="fileName">新文件路径</param>
    void SaveAs(string fileName);

    /// <summary>
    /// 生成编译文件
    /// 对应 VBProject.MakeCompiledFile 方法
    /// </summary>
    void MakeCompiledFile();
    #endregion
}
