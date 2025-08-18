//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
public interface IExcelProtectedViewWindows : IDisposable, IEnumerable<IExcelProtectedViewWindow>
{
    /// <summary>
    /// 获取受保护视图窗口集合中的窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取受保护视图窗口（索引从1开始）
    /// </summary>
    /// <param name="index">窗口索引</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow this[int index] { get; }

    /// <summary>
    /// 根据窗口标题获取受保护视图窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow this[string caption] { get; }

    /// <summary>
    /// 打开文件到受保护视图
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="password">密码</param>
    /// <param name="readOnlyRecommended">是否推荐只读</param>
    /// <param name="editable">是否可编辑</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow Open(string filename, string password = null,
                                  bool readOnlyRecommended = false, bool editable = false);

    /// <summary>
    /// 根据文件名查找受保护视图窗口
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow FindByFilename(string filename);

    /// <summary>
    /// 根据窗口标题查找受保护视图窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow FindByCaption(string caption);


    /// <summary>
    /// 获取父级应用程序
    /// </summary>
    IExcelApplication Parent { get; }

    /// <summary>
    /// 获取活动的受保护视图窗口
    /// </summary>
    IExcelProtectedViewWindow ActiveProtectedViewWindow { get; }

    /// <summary>
    /// 获取可见的受保护视图窗口
    /// </summary>
    IEnumerable<IExcelProtectedViewWindow> VisibleWindows { get; }

    /// <summary>
    /// 获取最大化状态的受保护视图窗口
    /// </summary>
    IEnumerable<IExcelProtectedViewWindow> MaximizedWindows { get; }

    /// <summary>
    /// 获取最小化状态的受保护视图窗口
    /// </summary>
    IEnumerable<IExcelProtectedViewWindow> MinimizedWindows { get; }

    /// <summary>
    /// 获取指定状态的窗口
    /// </summary>
    /// <param name="state">窗口状态</param>
    /// <returns>受保护视图窗口枚举</returns>
    IEnumerable<IExcelProtectedViewWindow> GetWindowsByState(XlProtectedViewWindowState state);
}