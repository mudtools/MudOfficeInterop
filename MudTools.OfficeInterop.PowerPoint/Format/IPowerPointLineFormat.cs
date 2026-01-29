//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


using System;
using System.Runtime.InteropServices;

/// <summary>
/// 表示 PowerPoint 形状的线条格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointLineFormat : IDisposable
{
    /// <summary>
    /// 获取创建此线条格式设置的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此线条格式设置的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此线条格式设置的父对象。
    /// </summary>
    /// <value>表示此线条格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置线条的背景颜色。
    /// </summary>
    /// <value>表示背景颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? BackColor { get; set; }

    /// <summary>
    /// 获取或设置线条起始箭头头的长度。
    /// </summary>
    /// <value>表示起始箭头头长度的 <see cref="MsoArrowheadLength"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条起始箭头头的样式。
    /// </summary>
    /// <value>表示起始箭头头样式的 <see cref="MsoArrowheadStyle"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条起始箭头头的宽度。
    /// </summary>
    /// <value>表示起始箭头头宽度的 <see cref="MsoArrowheadWidth"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth BeginArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的虚线样式。
    /// </summary>
    /// <value>表示虚线样式的 <see cref="MsoLineDashStyle"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineDashStyle DashStyle { get; set; }

    /// <summary>
    /// 获取或设置线条结束箭头头的长度。
    /// </summary>
    /// <value>表示结束箭头头长度的 <see cref="MsoArrowheadLength"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条结束箭头头的样式。
    /// </summary>
    /// <value>表示结束箭头头样式的 <see cref="MsoArrowheadStyle"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条结束箭头头的宽度。
    /// </summary>
    /// <value>表示结束箭头头宽度的 <see cref="MsoArrowheadWidth"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的前景颜色。
    /// </summary>
    /// <value>表示前景颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? ForeColor { get; set; }

    /// <summary>
    /// 获取或设置线条的图案样式。
    /// </summary>
    /// <value>表示图案样式的 <see cref="MsoPatternType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPatternType Pattern { get; set; }

    /// <summary>
    /// 获取或设置线条的样式。
    /// </summary>
    /// <value>表示线条样式的 <see cref="MsoLineStyle"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取或设置线条的透明度。
    /// </summary>
    /// <value>表示透明度值的浮点数（0.0 到 1.0）。</value>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示线条是否可见。
    /// </summary>
    /// <value>指示是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置线条的粗细（以磅为单位）。
    /// </summary>
    /// <value>表示线条粗细的浮点数。</value>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用内嵌笔。
    /// </summary>
    /// <value>指示是否使用内嵌笔的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool InsetPen { get; set; }
}