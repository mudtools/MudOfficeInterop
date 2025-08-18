//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel RecentFiles 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.RecentFiles 的安全访问和操作
/// </summary>
public interface IExcelRecentFiles : IEnumerable<IExcelRecentFile>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取最近使用的文件集合中的文件数量
    /// 对应 RecentFiles.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的最近使用文件对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">文件索引（从1开始）</param>
    /// <returns>最近使用文件对象</returns>
    IExcelRecentFile this[int index] { get; }

    /// <summary>
    /// 获取最近使用的文件集合所在的父对象（通常是 Application）
    /// 对应 RecentFiles.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取最近使用的文件集合所在的Application对象
    /// 对应 RecentFiles.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置最近使用的文件列表的最大容量
    /// 对应 Application.RecentFiles.Maximum 属性
    /// </summary>
    int Maximum { get; set; }
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找最近使用的文件
    /// </summary>
    /// <param name="name">文件名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的文件数组</returns>
    IExcelRecentFile[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据路径查找最近使用的文件
    /// </summary>
    /// <param name="path">文件路径</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的文件数组</returns>
    IExcelRecentFile[] FindByPath(string path, bool matchCase = false);

    /// <summary>
    /// 根据扩展名查找最近使用的文件
    /// </summary>
    /// <param name="extension">文件扩展名 (例如 ".xlsx", ".xls")</param>
    /// <returns>匹配的文件数组</returns>
    IExcelRecentFile[] FindByExtension(string extension);

    /// <summary>
    /// 获取最近访问的文件
    /// </summary>
    /// <param name="count">返回文件的数量</param>
    /// <returns>最近访问的文件数组</returns>
    IExcelRecentFile[] GetMostRecent(int count = 5);

    /// <summary>
    /// 获取最久未访问的文件
    /// </summary>
    /// <param name="count">返回文件的数量</param>
    /// <returns>最久未访问的文件数组</returns>
    IExcelRecentFile[] GetLeastRecent(int count = 5);
    #endregion

    #region 操作方法
    /// <summary>
    /// 清除所有最近使用的文件记录
    /// 对应 RecentFiles.Delete 方法 (对每个项目)
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的最近使用文件记录
    /// </summary>
    /// <param name="index">要删除的文件索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的最近使用文件对象
    /// </summary>
    /// <param name="file">要删除的文件对象</param>
    void Delete(IExcelRecentFile file);

    /// <summary>
    /// 批量删除最近使用文件记录
    /// </summary>
    /// <param name="indices">要删除的文件索引数组</param>
    void DeleteRange(int[] indices);
    #endregion

}
