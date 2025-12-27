//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel LineFormat 对象的二次封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelLineFormat : IOfficeObject<IExcelLineFormat>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取线条所在的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取线条对象所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }
    #endregion

    // 颜色相关
    /// <summary>
    /// 获取线条的前景色格式
    /// </summary>
    IExcelColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取线条的背景色格式
    /// </summary>
    IExcelColorFormat? BackColor { get; }

    /// <summary>
    /// 获取或设置线条的透明度
    /// </summary>
    float Transparency { get; set; }

    // 线型与样式
    /// <summary>
    /// 获取或设置线条的虚线样式（别名属性，实际对应DashStyle）
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取或设置线条的实际虚线样式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineDashStyle DashStyle { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头的长度
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头的样式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头的宽度
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth BeginArrowheadWidth { get; set; }

    // 终点箭头
    /// <summary>
    /// 获取或设置线条终点箭头的长度
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条终点箭头的样式
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条终点箭头的宽度
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的填充图案类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPatternType Pattern { get; set; }

    /// <summary>
    /// 获取或设置线条粗细
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置线条是否可见
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置是否使用内嵌画笔绘制线条
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool InsetPen { get; set; }
}