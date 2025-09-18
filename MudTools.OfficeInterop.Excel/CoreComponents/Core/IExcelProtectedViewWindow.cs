//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel受保护视图窗口的接口，提供对受保护视图窗口的各种操作和属性访问
/// </summary>
public interface IExcelProtectedViewWindow : IExcelCommonWindow, IDisposable
{
    /// <summary>
    /// 获取与受保护视图窗口关联的工作簿对象
    /// </summary>
    IExcelWorkbook? Workbook { get; }

    /// <summary>
    /// 获取受保护视图窗口的状态
    /// </summary>
    XlProtectedViewWindowState WindowState { get; set; }

    /// <summary>
    /// 获取窗口的文件路径
    /// </summary>
    string SourcePath { get; }

    /// <summary>
    /// 获取窗口的文件名
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 编辑受保护视图中的工作簿
    /// </summary>
    /// <returns>编辑后的工作簿对象</returns>
    IExcelWorkbook Edit();

    /// <summary>
    /// 最大化受保护视图窗口
    /// </summary>
    void Maximize();

    /// <summary>
    /// 最小化受保护视图窗口
    /// </summary>
    void Minimize();

    /// <summary>
    /// 恢复受保护视图窗口到正常大小
    /// </summary>
    void Restore();

    /// <summary>
    /// 移动受保护视图窗口到指定位置
    /// </summary>
    /// <param name="left">新左侧位置</param>
    /// <param name="top">新顶部位置</param>
    void Move(int left, int top);

    /// <summary>
    /// 调整受保护视图窗口大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    void Resize(int width, int height);
}