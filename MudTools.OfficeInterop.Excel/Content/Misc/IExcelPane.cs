//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Microsoft.Office.Interop.Excel.Pane 实现的二次封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPane : IOfficeObject<IExcelPane, MsExcel.Pane>, IDisposable
{
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    /// <summary>
    /// 获取窗格索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取窗格的可视区域范围
    /// </summary>
    IExcelRange? VisibleRange { get; }

    /// <summary>
    /// 获取窗格的滚动列位置
    /// </summary>
    int ScrollColumn { get; set; }

    /// <summary>
    /// 获取窗格的滚动行位置
    /// </summary>
    int ScrollRow { get; set; }

    /// <summary>
    /// 激活当前窗格
    /// </summary>
    void Activate();

    /// <summary>
    /// 将指定矩形区域滚动到窗格视图中
    /// </summary>
    /// <param name="left">矩形区域左上角的水平坐标（单位：磅）</param>
    /// <param name="top">矩形区域左上角的垂直坐标（单位：磅）</param>
    /// <param name="width">矩形区域的宽度（单位：磅）</param>
    /// <param name="height">矩形区域的高度（单位：磅）</param>
    /// <param name="start">是否将区域滚动到视图的起始位置</param>
    void ScrollIntoView(int left, int top, int width, int height, bool start = true);

    /// <summary>
    /// 将点单位转换为屏幕像素的水平坐标
    /// </summary>
    /// <param name="points">点单位坐标值</param>
    /// <returns>对应的屏幕像素X坐标</returns>
    int? PointsToScreenPixelsX(int points);

    /// <summary>
    /// 将点单位转换为屏幕像素的垂直坐标
    /// </summary>
    /// <param name="points">点单位坐标值</param>
    /// <returns>对应的屏幕像素Y坐标</returns>
    int? PointsToScreenPixelsY(int points);

    /// <summary>
    /// 执行大幅度滚动操作
    /// </summary>
    /// <param name="down">向下滚动的行数</param>
    /// <param name="up">向上滚动的行数</param>
    /// <param name="toRight">向右滚动的列数</param>
    /// <param name="toLeft">向左滚动的列数</param>
    /// <returns>操作结果对象</returns>
    object? LargeScroll(int down, int up, int toRight, int toLeft);

    /// <summary>
    /// 执行小幅度滚动操作
    /// </summary>
    /// <param name="down">向下滚动的行数</param>
    /// <param name="up">向上滚动的行数</param>
    /// <param name="toRight">向右滚动的列数</param>
    /// <param name="toLeft">向左滚动的列数</param>
    /// <returns>操作结果对象</returns>
    object? SmallScroll(int down, int up, int toRight, int toLeft);
}