//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE VBProjects 集合对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBProjects 的安全访问和操作
/// </summary>
public interface IVbeVBProjects : IEnumerable<IVbeVBProject>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 VB 项目集合中的项目数量
    /// 对应 VBProjects.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的 VB 项目对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">项目索引（从1开始）</param>
    /// <returns>VB 项目对象</returns>
    IVbeVBProject this[int index] { get; }

    /// <summary>
    /// 获取 VB 项目集合所在的父对象（通常是 VBE）
    /// 对应 VBProjects.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 VB 项目集合所在的Application对象（VBE 对象）
    /// 对应 VBProjects.Application 属性
    /// </summary>
    IVbeApplication Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的 VB 项目
    /// 对应 VBProjects.Add 方法
    /// </summary>
    /// <param name="projectType">项目类型</param>
    /// <param name="projectName">项目名称</param>
    /// <returns>新创建的 VB 项目对象</returns>
    IVbeVBProject Add(vbext_ProjectType projectType, string projectName = "");

    /// < <summary>
    /// 打开一个现有的 VB 项目文件
    /// 对应 VBProjects.Open 方法
    /// </summary>
    /// <param name="fileName">项目文件路径</param>
    /// <returns>打开的 VB 项目对象</returns>
    IVbeVBProject Open(string fileName);

    /// <summary>
    /// 基于模板创建 VB 项目 (概念性，VBA 通常不直接支持)
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="projectName">新项目名称</param>
    /// <returns>新创建的 VB 项目对象</returns>
    IVbeVBProject CreateFromTemplate(string templatePath, string projectName = "");
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找 VB 项目
    /// </summary>
    /// <param name="name">项目名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的 VB 项目数组</returns>
    IVbeVBProject[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据类型查找 VB 项目
    /// </summary>
    /// <param name="projectType">项目类型</param>
    /// <returns>匹配的 VB 项目数组</returns>
    IVbeVBProject[] FindByType(vbext_ProjectType projectType);

    /// <summary>
    /// 根据路径查找 VB 项目
    /// </summary>
    /// <param name="path">项目路径</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的 VB 项目数组</returns>
    IVbeVBProject[] FindByPath(string path, bool matchCase = false);

    /// <summary>
    /// 获取所有标准 EXE 项目
    /// </summary>
    /// <returns>标准 EXE 项目数组</returns>
    IVbeVBProject[] GetStandardExeProjects();

    /// <summary>
    /// 获取所有 DLL 项目
    /// </summary>
    /// <returns>DLL 项目数组</returns>
    IVbeVBProject[] GetDllProjects();

    /// <summary>
    /// 获取所有受保护的项目
    /// </summary>
    /// <returns>受保护项目数组</returns>
    IVbeVBProject[] GetProtectedProjects();
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除所有 VB 项目
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的 VB 项目
    /// </summary>
    /// <param name="index">要删除的项目索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的 VB 项目
    /// </summary>
    /// <param name="name">要删除的项目名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的 VB 项目对象
    /// </summary>
    /// <param name="project">要删除的 VB 项目对象</param>
    void Delete(IVbeVBProject project);

    /// <summary>
    /// 批量删除 VB 项目
    /// </summary>
    /// <param name="indices">要删除的项目索引数组 (建议降序排列)</param>
    void DeleteRange(int[] indices);
    #endregion

    #region 导出和导入 (概念性)
    /// <summary>
    /// 导出所有 VB 项目到文件夹
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的项目数量</returns>
    int ExportToFolder(string folderPath, string prefix = "project_");

    /// <summary>
    /// 从文件夹导入 VB 项目
    /// </summary>
    /// <param name="folderPath">导入文件夹路径</param>
    /// <returns>成功导入的项目数量</returns>
    int ImportFromFolder(string folderPath);

    #endregion

    #region 高级功能
    /// <summary>
    /// 获取活动的 VB 项目
    /// </summary>
    /// <returns>活动 VB 项目对象</returns>
    IVbeVBProject ActiveProject { get; }

    /// <summary>
    /// 编译所有 VB 项目
    /// </summary>
    void CompileAll();
    #endregion
}
