//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


using System;
using System.Runtime.InteropServices;

/// <summary>
/// 表示 PowerPoint 演示文稿中的设计模板。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointDesign : IOfficeObject<IPowerPointDesign, MsPowerPoint.Design>, IDisposable
{
    /// <summary>
    /// 获取创建此设计模板的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此设计模板的父对象。
    /// </summary>
    /// <value>表示此设计模板父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此设计模板的幻灯片母版。
    /// </summary>
    /// <value>表示幻灯片母版的 <see cref="IPowerPointMaster"/> 对象。</value>
    IPowerPointMaster? SlideMaster { get; }

    /// <summary>
    /// 获取此设计模板的标题母版。
    /// </summary>
    /// <value>表示标题母版的 <see cref="IPowerPointMaster"/> 对象。</value>
    IPowerPointMaster? TitleMaster { get; }

    /// <summary>
    /// 获取一个值，指示此设计模板是否有标题母版。
    /// </summary>
    /// <value>指示是否有标题母版的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTitleMaster { get; }

    /// <summary>
    /// 为此设计模板添加标题母版。
    /// </summary>
    /// <returns>新添加的标题母版。</returns>
    IPowerPointMaster? AddTitleMaster();

    /// <summary>
    /// 获取此设计模板在集合中的索引。
    /// </summary>
    /// <value>表示索引的整数值。</value>
    int Index { get; }

    /// <summary>
    /// 获取或设置此设计模板的名称。
    /// </summary>
    /// <value>表示设计模板名称的字符串。</value>
    string? Name { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示此设计模板是否被保留。
    /// </summary>
    /// <value>指示设计模板是否被保留的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Preserved { get; set; }

    /// <summary>
    /// 将此设计模板移动到指定位置。
    /// </summary>
    /// <param name="toPos">要移动到的目标位置索引。</param>
    void MoveTo(int toPos);

    /// <summary>
    /// 删除此设计模板。
    /// </summary>
    void Delete();
}