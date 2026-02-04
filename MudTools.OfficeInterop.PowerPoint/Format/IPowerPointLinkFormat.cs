//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

using System;

/// <summary>
/// 表示链接格式，用于管理 PowerPoint 中链接对象（如链接的 OLE 对象）的属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointLinkFormat : IOfficeObject<IPowerPointLinkFormat, MsPowerPoint.LinkFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此对象的 PowerPoint 应用程序。
    /// </summary>
    /// <value>PowerPoint 应用程序对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此链接格式的父对象。
    /// </summary>
    /// <value>父对象。</value>
    object Parent { get; }

    /// <summary>
    /// 获取或设置链接对象的完整源文件路径和名称。
    /// </summary>
    /// <value>源文件的完整路径字符串。</value>
    string SourceFullName { get; set; }

    /// <summary>
    /// 获取或设置链接对象的自动更新选项。
    /// </summary>
    /// <value>更新选项的枚举值。</value>
    PpUpdateOption AutoUpdate { get; set; }

    /// <summary>
    /// 手动更新链接对象，从源文件获取最新数据。
    /// </summary>
    void Update();

    /// <summary>
    /// 断开与源文件的链接，将对象转换为静态副本。
    /// </summary>
    void BreakLink();
}