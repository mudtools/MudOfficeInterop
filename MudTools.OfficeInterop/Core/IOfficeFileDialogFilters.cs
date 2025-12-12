//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// Office FileDialogFilters 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Core.FileDialogFilters 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
[ItemIndex]
public interface IOfficeFileDialogFilters : IEnumerable<IOfficeFileDialogFilter>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取文件对话框过滤器集合中的过滤器数量
    /// 对应 FileDialogFilters.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的文件对话框过滤器对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">过滤器索引（从1开始）</param>
    /// <returns>文件对话框过滤器对象</returns>
    IOfficeFileDialogFilter this[int index] { get; }

    /// <summary>
    /// 获取过滤器集合所在的父对象（通常是 FileDialog）
    /// 对应 FileDialogFilters.Parent 属性
    /// </summary>
    object Parent { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的文件过滤器
    /// 对应 FileDialogFilters.Add 方法
    /// </summary>
    /// <param name="description">过滤器描述 (例如 "Text Files")</param>
    /// <param name="extensions">过滤器扩展名 (例如 "*.txt" 或 "*.txt;*.csv")</param>
    /// <param name="position">插入位置 (从1开始，默认添加到末尾)</param>
    /// <returns>新创建的文件对话框过滤器对象</returns>
    IOfficeFileDialogFilter Add(string description, string extensions, int position = -1);
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除集合中的所有过滤器
    /// 对应 FileDialogFilters.Clear 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定的过滤器对象
    /// </summary>
    /// <param name="filter">要删除的过滤器对象</param>
    void Delete(IOfficeFileDialogFilter filter);
    #endregion

}
