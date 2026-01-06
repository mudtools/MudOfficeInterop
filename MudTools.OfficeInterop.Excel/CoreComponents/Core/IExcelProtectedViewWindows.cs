//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示Excel受保护视图窗口集合的接口
/// 继承自IDisposable和IEnumerable接口，支持对受保护视图窗口进行枚举和资源释放
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelProtectedViewWindows : IDisposable, IOfficeObject<IExcelProtectedViewWindows, MsExcel.ProtectedViewWindows>, IEnumerable<IExcelProtectedViewWindow?>
{
    /// <summary>
    /// 获取对象的父对象（通常是 Application）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取受保护视图窗口集合中的窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取受保护视图窗口（索引从1开始）
    /// </summary>
    /// <param name="index">窗口索引</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow? this[int index] { get; }

    /// <summary>
    /// 根据窗口标题获取受保护视图窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow? this[string caption] { get; }

    /// <summary>
    /// 打开文件到受保护视图
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="password">密码</param>
    /// <param name="addToMru"></param>
    /// <param name="repairMode"></param>
    /// <returns>受保护视图窗口对象</returns>
    IExcelProtectedViewWindow? Open(string filename, string? password = null,
                                  bool? addToMru = false, bool? repairMode = false);


}