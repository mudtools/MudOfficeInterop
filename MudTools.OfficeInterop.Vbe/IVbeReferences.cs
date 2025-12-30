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
[ComCollectionWrap(ComNamespace = "MsVb"), ItemIndex]
public interface IVbeReferences : IEnumerable<IVbeReference?>, IOfficeObject<IVbeReferences>, IDisposable
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
    IVbeReference? this[int index] { get; }

    /// <summary>
    /// 获取引用集合所在的父对象（通常是 VBProject）
    /// 对应 References.Parent 属性
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取 VB 项目的 VBE 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }
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
    IVbeReference? AddFromGuid(string guid, int major, int minor);

    /// <summary>
    /// 向集合中添加新的引用（通过文件路径）
    /// 对应 References.AddFromFile 方法
    /// </summary>
    /// <param name="fileName">类型库文件路径</param>
    /// <returns>新添加的引用对象</returns>
    IVbeReference? AddFromFile(string fileName);

    /// <summary>
    /// 移除指定的引用
    /// </summary>
    /// <param name="reference"></param>
    void Remove(IVbeReference reference);
    #endregion
}
