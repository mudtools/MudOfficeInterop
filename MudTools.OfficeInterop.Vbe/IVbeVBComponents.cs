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
public interface IVbeVBComponents : IEnumerable<IVbeVBComponent>, IDisposable
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
    IVbeVBComponent this[int index] { get; }

    /// <summary>
    /// 获取指定名称的 VB 组件对象
    /// </summary>
    /// <param name="name">组件名称</param>
    /// <returns>VB 组件对象</returns>
    IVbeVBComponent this[string name] { get; }

    /// <summary>
    /// 获取 VB 组件集合所在的父对象（通常是 VBProject）
    /// 对应 VBComponents.Parent 属性
    /// </summary>
    object? Parent { get; }

    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的 VB 组件
    /// 对应 VBComponents.Add 方法
    /// </summary>
    /// <param name="componentType">组件类型</param>
    /// <param name="name">组件名称 (可选)</param>
    /// <returns>新创建的 VB 组件对象</returns>
    IVbeVBComponent Add(vbext_ComponentType componentType, string name = "");

    /// <summary>
    /// 从文件导入 VB 组件
    /// 对应 VBComponents.Import 方法
    /// </summary>
    /// <param name="fileName">组件文件路径</param>
    /// <returns>导入的 VB 组件对象</returns>
    IVbeVBComponent Import(string fileName);

    /// <summary>
    /// 基于模板创建 VB 组件 (概念性，VBA 通常不直接支持)
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="name">新组件名称</param>
    /// <returns>新创建的 VB 组件对象</returns>
    IVbeVBComponent CreateFromTemplate(string templatePath, string name = "");
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找 VB 组件
    /// </summary>
    /// <param name="name">组件名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的 VB 组件数组</returns>
    IVbeVBComponent[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据类型查找 VB 组件
    /// </summary>
    /// <param name="componentType">组件类型</param>
    /// <returns>匹配的 VB 组件数组</returns>
    IVbeVBComponent[] FindByType(vbext_ComponentType componentType); // componentType 使用 int 代表 vbext_ComponentType

    /// <summary>
    /// 获取所有标准模块
    /// </summary>
    /// <returns>标准模块数组</returns>
    IVbeVBComponent[] GetStandardModules();

    /// <summary>
    /// 获取所有类模块
    /// </summary>
    /// <returns>类模块数组</returns>
    IVbeVBComponent[] GetClassModules();

    /// <summary>
    /// 获取所有用户窗体
    /// </summary>
    /// <returns>用户窗体数组</returns>
    IVbeVBComponent[] GetUserForms();

    /// <summary>
    /// 获取所有文档模块（特定于主机，如 Excel.Worksheet, Excel.Workbook）
    /// </summary>
    /// <returns>文档模块数组</returns>
    IVbeVBComponent[] GetDocumentModules();
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除所有 VB 组件
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的 VB 组件
    /// </summary>
    /// <param name="index">要删除的组件索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的 VB 组件
    /// </summary>
    /// <param name="name">要删除的组件名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的 VB 组件对象
    /// </summary>
    /// <param name="component">要删除的 VB 组件对象</param>
    void Delete(IVbeVBComponent component);

    /// <summary>
    /// 批量删除 VB 组件
    /// </summary>
    /// <param name="indices">要删除的组件索引数组</param>
    void DeleteRange(int[] indices);
    #endregion

    #region 导出和导入
    /// <summary>
    /// 导出所有 VB 组件到文件夹
    /// </summary>
    /// <param name="folderPath">导出文件夹路径</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>成功导出的组件数量</returns>
    int ExportToFolder(string folderPath, string prefix = "component_");

    /// <summary>
    /// 从文件夹导入 VB 组件
    /// </summary>
    /// <param name="folderPath">导入文件夹路径</param>
    /// <returns>成功导入的组件数量</returns>
    int ImportFromFolder(string folderPath);
    #endregion
}
