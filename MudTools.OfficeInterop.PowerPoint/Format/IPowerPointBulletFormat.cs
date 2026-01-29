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
/// 表示 PowerPoint 段落中项目符号的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointBulletFormat : IDisposable
{
    /// <summary>
    /// 获取创建此项目符号格式设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此项目符号格式设置的父对象。
    /// </summary>
    /// <value>表示此项目符号格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示项目符号是否可见。
    /// </summary>
    /// <value>指示项目符号是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置用作项目符号的字符的 Unicode 编码。
    /// </summary>
    /// <value>表示项目符号字符的整数值。</value>
    int Character { get; set; }

    /// <summary>
    /// 获取或设置项目符号相对于文本的相对大小。
    /// </summary>
    /// <value>表示相对大小的浮点数。</value>
    float RelativeSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示项目符号是否使用文本颜色。
    /// </summary>
    /// <value>指示是否使用文本颜色的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseTextColor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示项目符号是否使用文本字体。
    /// </summary>
    /// <value>指示是否使用文本字体的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseTextFont { get; set; }

    /// <summary>
    /// 获取项目符号的字体设置。
    /// </summary>
    /// <value>表示项目符号字体的 <see cref="IPowerPointFont"/> 对象。</value>
    IPowerPointFont? Font { get; }

    /// <summary>
    /// 获取或设置项目符号的类型。
    /// </summary>
    /// <value>表示项目符号类型的 <see cref="PpBulletType"/> 枚举值。</value>
    PpBulletType Type { get; set; }

    /// <summary>
    /// 获取或设置编号项目符号的样式。
    /// </summary>
    /// <value>表示编号项目符号样式的 <see cref="PpNumberedBulletStyle"/> 枚举值。</value>
    PpNumberedBulletStyle Style { get; set; }

    /// <summary>
    /// 获取或设置编号项目符号的起始值。
    /// </summary>
    /// <value>表示编号起始值的整数值。</value>
    int StartValue { get; set; }

    /// <summary>
    /// 将图片设置为项目符号。
    /// </summary>
    /// <param name="picture">要设置为项目符号的图片文件路径。</param>
    void Picture(string picture);

    /// <summary>
    /// 获取当前段落的项目符号编号。
    /// </summary>
    /// <value>表示项目符号编号的整数值。</value>
    int Number { get; }
}