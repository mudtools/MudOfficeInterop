//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示标注格式的设置和操作。
/// 该接口提供对标注线的各种属性（如角度、长度、类型等）进行配置和查询的功能。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCalloutFormat : IDisposable
{
    /// <summary>
    /// 获取创建此对象的应用程序对象。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此对象的父对象。
    /// </summary>
    /// <value>表示父对象的对象。</value>
    object Parent { get; }

    /// <summary>
    /// 将标注线设置为自动长度。
    /// </summary>
    void AutomaticLength();

    /// <summary>
    /// 设置标注线的自定义下落距离。
    /// </summary>
    /// <param name="drop">下落距离的数值。</param>
    void CustomDrop(float drop);

    /// <summary>
    /// 设置标注线的自定义长度。
    /// </summary>
    /// <param name="length">长度的数值。</param>
    void CustomLength(float length);

    /// <summary>
    /// 使用预设的下落类型设置标注线的下落距离。
    /// </summary>
    /// <param name="dropType">预设的下落类型。</param>
    void PresetDrop([ComNamespace("MsCore")] MsoCalloutDropType dropType);

    /// <summary>
    /// 获取或设置一个值，指示标注线是否具有强调线。
    /// </summary>
    /// <value>如果标注线具有强调线，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Accent { get; set; }

    /// <summary>
    /// 获取或设置标注线的角度类型。
    /// </summary>
    /// <value>标注线的角度类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutAngleType Angle { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当标注线移动时，标注线是否自动连接到标注框。
    /// </summary>
    /// <value>如果标注线自动连接，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoAttach { get; set; }

    /// <summary>
    /// 获取一个值，指示标注线长度是否为自动。
    /// </summary>
    /// <value>如果标注线长度为自动，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoLength { get; }

    /// <summary>
    /// 获取或设置一个值，指示标注线是否具有边框。
    /// </summary>
    /// <value>如果标注线具有边框，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Border { get; set; }

    /// <summary>
    /// 获取标注线的下落距离（从标注线连接到标注框的点到标注线开始的位置）。
    /// </summary>
    /// <value>下落距离的数值。</value>
    float Drop { get; }

    /// <summary>
    /// 获取标注线的下落类型。
    /// </summary>
    /// <value>标注线的下落类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutDropType DropType { get; }

    /// <summary>
    /// 获取或设置标注线与标注框之间的间隙距离。
    /// </summary>
    /// <value>间隙距离的数值。</value>
    float Gap { get; set; }

    /// <summary>
    /// 获取标注线的长度（从标注线连接到标注框的点到标注线指向的对象的距离）。
    /// </summary>
    /// <value>长度的数值。</value>
    float Length { get; }

    /// <summary>
    /// 获取或设置标注线的类型。
    /// </summary>
    /// <value>标注线的类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutType Type { get; set; }
}