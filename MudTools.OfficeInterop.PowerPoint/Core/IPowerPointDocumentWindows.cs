//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint DocumentWindows 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.DocumentWindows 的安全访问和操作
/// </summary>
public interface IPowerPointDocumentWindows : IEnumerable<IPowerPointDocumentWindow>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取窗口集合中的窗口数量
    /// 对应 DocumentWindows.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的窗口对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">窗口索引（从1开始）</param>
    /// <returns>窗口对象</returns>
    IPowerPointDocumentWindow this[int index] { get; }

    /// <summary>
    /// 获取窗口集合所在的父对象（通常是 Application）
    /// 对应 DocumentWindows.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取窗口集合所在的Application对象
    /// 对应 DocumentWindows.Application 属性
    /// </summary>
    IPowerPointApplication Application { get; }
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据标题查找窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的窗口数组</returns>
    IPowerPointDocumentWindow[] FindByCaption(string caption, bool matchCase = false);

    /// <summary>
    /// 根据关联的演示文稿名称查找窗口
    /// </summary>
    /// <param name="presentationName">演示文稿名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的窗口数组</returns>
    IPowerPointDocumentWindow[] FindByPresentationName(string presentationName, bool matchCase = false);

    /// <summary>
    /// 获取活动的窗口
    /// </summary>
    /// <returns>活动窗口对象</returns>
    IPowerPointDocumentWindow GetActiveWindow();

    /// <summary>
    /// 获取可见的窗口
    /// </summary>
    /// <returns>可见窗口数组</returns>
    IPowerPointDocumentWindow[] GetVisibleWindows();

    #endregion

    #region 操作方法
    /// <summary>
    /// 关闭所有窗口
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除/关闭指定索引的窗口
    /// </summary>
    /// <param name="index">要关闭的窗口索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除/关闭指定的窗口对象
    /// </summary>
    /// <param name="window">要关闭的窗口对象</param>
    void Delete(IPowerPointDocumentWindow window);

    /// <summary>
    /// 批量删除/关闭窗口
    /// </summary>
    /// <param name="indices">要关闭的窗口索引数组 (建议降序排列)</param>
    void DeleteRange(int[] indices);

    #endregion
}
