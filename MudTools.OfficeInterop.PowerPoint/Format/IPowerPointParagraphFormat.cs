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
/// 表示 PowerPoint 文本框中段落的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointParagraphFormat : IOfficeObject<IPowerPointParagraphFormat, MsPowerPoint.ParagraphFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此段落格式设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此段落格式设置的父对象。
    /// </summary>
    /// <value>表示此段落格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置段落的对齐方式。
    /// </summary>
    /// <value>表示段落对齐方式的 <see cref="PpParagraphAlignment"/> 枚举值。</value>
    PpParagraphAlignment Alignment { get; set; }

    /// <summary>
    /// 获取段落的项目符号格式设置。
    /// </summary>
    /// <value>表示项目符号格式的 <see cref="IPowerPointBulletFormat"/> 对象。</value>
    IPowerPointBulletFormat? Bullet { get; }

    /// <summary>
    /// 获取或设置一个值，指示段前间距是否使用行距规则。
    /// </summary>
    /// <value>指示是否使用行距规则的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleBefore { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示段后间距是否使用行距规则。
    /// </summary>
    /// <value>指示是否使用行距规则的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleAfter { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示段内行间距是否使用行距规则。
    /// </summary>
    /// <value>指示是否使用行距规则的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LineRuleWithin { get; set; }

    /// <summary>
    /// 获取或设置段前间距（以磅为单位）。
    /// </summary>
    /// <value>表示段前间距的浮点数。</value>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置段后间距（以磅为单位）。
    /// </summary>
    /// <value>表示段后间距的浮点数。</value>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置段内行间距（以磅为单位）。
    /// </summary>
    /// <value>表示段内行间距的浮点数。</value>
    float SpaceWithin { get; set; }

    /// <summary>
    /// 获取或设置文本的基线对齐方式。
    /// </summary>
    /// <value>表示基线对齐方式的 <see cref="PpBaselineAlignment"/> 枚举值。</value>
    PpBaselineAlignment BaseLineAlignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用远东换行控制。
    /// </summary>
    /// <value>指示是否启用远东换行控制的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FarEastLineBreakControl { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用自动换行。
    /// </summary>
    /// <value>指示是否启用自动换行的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用悬挂标点。
    /// </summary>
    /// <value>指示是否启用悬挂标点的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HangingPunctuation { get; set; }

    /// <summary>
    /// 获取或设置文本的方向。
    /// </summary>
    /// <value>表示文本方向的 <see cref="PpDirection"/> 枚举值。</value>
    PpDirection TextDirection { get; set; }
}