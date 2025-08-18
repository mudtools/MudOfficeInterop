//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE References 集合对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.References 的安全访问和操作
/// </summary>
public interface IVbeReferences : IEnumerable<IVbeReference>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取引用集合中的引用数量
    /// 对应 References.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的引用对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">引用索引（从1开始）</param>
    /// <returns>引用对象</returns>
    IVbeReference this[int index] { get; }

    /// <summary>
    /// 获取引用集合所在的父对象（通常是 VBProject）
    /// 对应 References.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取引用集合所在的Application对象（VBE 对象）
    /// 对应 References.Application 属性
    /// </summary>
    IVbeApplication Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的引用（通过 GUID）
    /// 对应 References.AddFromGuid 方法
    /// </summary>
    /// <param name="guid">类型库 GUID</param>
    /// <param name="major">主版本号</param>
    /// <param name="minor">次版本号</param>
    /// <returns>新添加的引用对象</returns>
    IVbeReference AddFromGuid(string guid, int major, int minor);

    /// <summary>
    /// 向集合中添加新的引用（通过文件路径）
    /// 对应 References.AddFromFile 方法
    /// </summary>
    /// <param name="fileName">类型库文件路径</param>
    /// <returns>新添加的引用对象</returns>
    IVbeReference AddFromFile(string fileName);

    /// <summary>
    /// 基于现有引用创建引用（概念性，通常通过添加实现）
    /// </summary>
    /// <param name="sourceReference">源引用对象</param>
    /// <returns>新创建的引用对象</returns>
    IVbeReference CopyFrom(IVbeReference sourceReference);
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找引用
    /// </summary>
    /// <param name="name">引用名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的引用数组</returns>
    IVbeReference[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据 GUID 查找引用
    /// </summary>
    /// <param name="guid">类型库 GUID</param>
    /// <returns>匹配的引用数组</returns>
    IVbeReference[] FindByGuid(string guid);

    /// <summary>
    /// 根据路径查找引用
    /// </summary>
    /// <param name="path">引用文件路径</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的引用数组</returns>
    IVbeReference[] FindByPath(string path, bool matchCase = false);

    /// <summary>
    /// 获取所有内置引用
    /// </summary>
    /// <returns>内置引用数组</returns>
    IVbeReference[] GetBuiltInReferences();

    /// <summary>
    /// 获取所有项目引用（相对于外部库）
    /// </summary>
    /// <returns>项目引用数组</returns>
    IVbeReference[] GetProjectReferences();

    /// <summary>
    /// 获取所有破损的引用
    /// </summary>
    /// <returns>破损引用数组</returns>
    IVbeReference[] GetBrokenReferences();

    /// <summary>
    /// 获取所有有效的引用
    /// </summary>
    /// <returns>有效引用数组</returns>
    IVbeReference[] GetValidReferences();
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除所有引用（危险操作！）
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的引用
    /// </summary>
    /// <param name="index">要删除的引用索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的引用
    /// </summary>
    /// <param name="name">要删除的引用名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的引用对象
    /// </summary>
    /// <param name="reference">要删除的引用对象</param>
    void Delete(IVbeReference reference);

    /// <summary>
    /// 批量删除引用
    /// </summary>
    /// <param name="indices">要删除的引用索引数组 (建议降序排列)</param>
    void DeleteRange(int[] indices);

    #endregion

    #region 导出和导入 (概念性)

    /// <summary>
    /// 从文件导入引用信息 (例如，从文本文件读取并尝试添加引用)
    /// </summary>
    /// <param name="filePath">导入文件路径</param>
    /// <returns>成功导入的引用数量</returns>
    int ImportFromFile(string filePath);
    #endregion

}
