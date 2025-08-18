//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 窗口集合接口
/// </summary>
public interface IWordWindows : IDisposable, IEnumerable<IWordWindow>
{
    /// <summary>
    /// 获取窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象（通常是 Application）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 根据索引获取窗口（从1开始）
    /// </summary>
    IWordWindow Item(int index);

    /// <summary>
    /// 根据窗口标题获取窗口
    /// </summary>
    IWordWindow Item(string caption);

    /// <summary>
    /// 创建新窗口
    /// </summary>
    /// <returns>新创建的窗口对象</returns>
    IWordWindow NewWindow();

    /// <summary>
    /// 获取活动窗口
    /// </summary>
    /// <returns>活动窗口对象</returns>
    IWordWindow GetActiveWindow();

    /// <summary>
    /// 根据条件查找窗口
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的窗口列表</returns>
    IEnumerable<IWordWindow> Find(Func<IWordWindow, bool> predicate);

    /// <summary>
    /// 按窗口标题排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的窗口列表</returns>
    IEnumerable<IWordWindow> OrderByCaption(bool ascending = true);

    /// <summary>
    /// 按窗口索引排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的窗口列表</returns>
    IEnumerable<IWordWindow> OrderByIndex(bool ascending = true);

    /// <summary>
    /// 刷新所有窗口
    /// </summary>
    void RefreshAll();

    /// <summary>
    /// 关闭所有窗口（除了指定的窗口）
    /// </summary>
    /// <param name="exceptWindow">要保留的窗口</param>
    void CloseAllExcept(IWordWindow exceptWindow = null);

    /// <summary>
    /// 激活所有窗口
    /// </summary>
    void ActivateAll();
}