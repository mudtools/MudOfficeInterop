
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.LineFormat 的接口，用于操作线条格式。
/// </summary>
public interface IWordLineFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置线条的前景颜色格式。
    /// </summary>
    IWordColorFormat ForeColor { get; }

    /// <summary>
    /// 获取或设置线条的背景颜色格式。
    /// </summary>
    IWordColorFormat BackColor { get; }

    /// <summary>
    /// 获取或设置线条的透明度（0.0到1.0之间）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置线条是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置线条的粗细（磅）。
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置线条样式。
    /// </summary>
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取或设置线条的虚线样式。
    /// </summary>
    MsoLineDashStyle DashStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的端点样式。
    /// </summary>
    MsoArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的起始箭头宽度。
    /// </summary>
    MsoArrowheadWidth BeginArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的起始箭头长度。
    /// </summary>
    MsoArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条的结束端点样式。
    /// </summary>
    MsoArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的结束箭头宽度。
    /// </summary>
    MsoArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的结束箭头长度。
    /// </summary>
    MsoArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置图案类型。
    /// </summary>
    MsoPatternType Pattern { get; set; }

    /// <summary>
    /// 设置纯色线条。
    /// </summary>
    /// <param name="color">RGB颜色值。</param>
    void Solid(int color);

    /// <summary>
    /// 设置虚线样式。
    /// </summary>
    /// <param name="dashStyle">虚线样式。</param>
    void SetDashStyle(MsoLineDashStyle dashStyle);

    /// <summary>
    /// 清除线条格式。
    /// </summary>
    void Clear();

    /// <summary>
    /// 复制线条格式到另一个对象。
    /// </summary>
    /// <param name="targetLine">目标线条格式对象。</param>
    void CopyTo(IWordLineFormat targetLine);

    /// <summary>
    /// 重置线条格式为默认值。
    /// </summary>
    void Reset();
}