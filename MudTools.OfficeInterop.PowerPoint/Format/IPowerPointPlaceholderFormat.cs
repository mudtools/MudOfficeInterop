//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 中占位符的格式信息。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPlaceholderFormat : IOfficeObject<IPowerPointPlaceholderFormat, MsPowerPoint.PlaceholderFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此占位符格式对象的 PowerPoint 应用程序实例。
    /// </summary>
    /// <returns>表示 PowerPoint 应用程序的对象。</returns>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此占位符格式对象的父对象。
    /// </summary>
    /// <returns>父对象。</returns>
    object Parent { get; }

    /// <summary>
    /// 获取占位符的类型。
    /// </summary>
    /// <returns>表示占位符类型的枚举值。</returns>
    PpPlaceholderType Type { get; }

    /// <summary>
    /// 获取或设置占位符的名称。
    /// </summary>
    /// <returns>占位符的名称。</returns>
    string Name { get; set; }

    /// <summary>
    /// 获取占位符所包含的形状类型。
    /// </summary>
    /// <returns>表示形状类型的枚举值。</returns>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeType ContainedType { get; }
}