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
public interface IVbeVBProject : IDisposable
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
    /// 获取 VB 项目所在的Application对象（VBE 对象）
    /// 对应 VBProject.Application 属性
    /// </summary>
    IVbeApplication Application { get; }

    /// <summary>
    /// 获取 VB 项目的完整路径（如果已保存）
    /// 对应 VBProject.FileName 属性
    /// </summary>
    string FileName { get; }

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

    #region 状态属性
    /// <summary>
    /// 获取 VB 项目是否已保存
    /// </summary>
    bool IsSaved { get; }

    /// <summary>
    /// 获取 VB 项目是否被保护
    /// </summary>
    bool IsProtected { get; }
    #endregion

    #region 核心对象集合
    /// <summary>
    /// 获取 VB 项目的组件集合
    /// 对应 VBProject.VBComponents 属性
    /// </summary>
    IVbeVBComponents VBComponents { get; }

    /// <summary>
    /// 获取 VB 项目的引用集合
    /// 对应 VBProject.References 属性
    /// </summary>
    IVbeReferences References { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 选择 VB 项目
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除 VB 项目（从其父 VBProjects 集合中移除）
    /// 对应 VBProjects.Remove 方法 (间接)
    /// </summary>
    void Delete();

    /// <summary>
    /// 保存 VB 项目
    /// </summary>
    void Save();

    /// <summary>
    /// 另存 VB 项目为
    /// 对应 VBProject.SaveAs 方法
    /// </summary>
    /// <param name="fileName">新文件路径</param>
    void SaveAs(string fileName);

    /// <summary>
    /// 刷新 VB 项目显示
    /// </summary>
    void Refresh();
    #endregion

    #region 项目操作
    /// <summary>
    /// 编译 VB 项目
    /// </summary>
    void Compile();

    /// <summary>
    /// 运行 VB 项目中的启动对象（如果定义）
    /// </summary>
    void Run();
    #endregion

    #region 引用管理
    /// <summary>
    /// 添加引用到 VB 项目
    /// 对应 References.AddFromGuid 或 AddFromFile 方法
    /// </summary>
    /// <param name="reference">引用对象 (GUID/文件路径/类型库)</param>
    /// <param name="guid">类型库 GUID (可选)</param>
    /// <param name="major">版本号 (可选)</param>
    /// <param name="minor">描述 (可选)</param>
    /// <returns>新添加的引用对象</returns>
    IVbeReference AddReference(object reference, string guid, int major, int minor);

    /// <summary>
    /// 移除 VB 项目中的引用
    /// 对应 References.Remove 方法
    /// </summary>
    /// <param name="reference">要移除的引用对象</param>
    void RemoveReference(IVbeReference reference);
    #endregion

    #region 导出和转换
    /// <summary>
    /// 获取 VB 项目的代码文本
    /// </summary>
    /// <returns>项目所有代码的文本</returns>
    string GetCodeText();
    #endregion

}
