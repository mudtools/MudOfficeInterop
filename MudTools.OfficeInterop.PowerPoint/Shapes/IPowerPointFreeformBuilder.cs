//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示自由形状构建器，用于通过添加节点来创建自定义的自由形状。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointFreeformBuilder : IOfficeObject<IPowerPointFreeformBuilder, MsPowerPoint.FreeformBuilder>, IDisposable
{
    /// <summary>
    /// 获取创建此自由形状构建器的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取自由形状构建器的父对象。
    /// </summary>
    /// <value>父对象，通常是形状集合或画布。</value>
    object? Parent { get; }

    /// <summary>
    /// 向自由形状添加节点。
    /// </summary>
    /// <param name="segmentType">要添加的线段类型。</param>
    /// <param name="editingType">节点的编辑类型。</param>
    /// <param name="x1">第一个控制点的X坐标（磅）。</param>
    /// <param name="y1">第一个控制点的Y坐标（磅）。</param>
    /// <param name="x2">第二个控制点的X坐标（磅），对于某些线段类型可选。</param>
    /// <param name="y2">第二个控制点的Y坐标（磅），对于某些线段类型可选。</param>
    /// <param name="x3">第三个控制点的X坐标（磅），对于某些线段类型可选。</param>
    /// <param name="y3">第三个控制点的Y坐标（磅），对于某些线段类型可选。</param>
    void AddNodes([ComNamespace("MsCore")] MsoSegmentType segmentType, [ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1, float x2 = 0f, float y2 = 0f, float x3 = 0f, float y3 = 0f);

    /// <summary>
    /// 将自由形状构建器转换为实际的形状对象。
    /// </summary>
    /// <returns>新创建的自由形状对象。</returns>
    IPowerPointShape? ConvertToShape();
}