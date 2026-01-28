//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 字体接口
/// </summary>
/// <summary>
/// 表示字体对象，用于设置和获取字体的各种属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointFont : IDisposable
{
    /// <summary>
    /// 获取字体所属的应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取字体的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取字体的颜色格式。
    /// </summary>
    IPowerPointColorFormat? Color { get; }

    /// <summary>
    /// 获取或设置字体是否加粗。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置字体是否倾斜。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置字体是否有阴影。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置字体是否有浮雕效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Emboss { get; set; }

    /// <summary>
    /// 获取或设置字体是否有下划线。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Underline { get; set; }

    /// <summary>
    /// 获取或设置字体是否为下标。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置字体是否为上标。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置基线的偏移量（以磅为单位）。
    /// </summary>
    float BaselineOffset { get; set; }

    /// <summary>
    /// 获取字体是否已嵌入文档中。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Embedded { get; }

    /// <summary>
    /// 获取字体是否可以嵌入文档中。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Embeddable { get; }

    /// <summary>
    /// 获取或设置字体大小（以磅为单位）。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置字体名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置远东地区字体名称。
    /// </summary>
    string NameFarEast { get; set; }

    /// <summary>
    /// 获取或设置ASCII字体名称。
    /// </summary>
    string NameAscii { get; set; }

    /// <summary>
    /// 获取或设置是否自动旋转数字。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoRotateNumbers { get; set; }

    /// <summary>
    /// 获取或设置其他字体名称。
    /// </summary>
    string NameOther { get; set; }

    /// <summary>
    /// 获取或设置复杂脚本字体名称。
    /// </summary>
    string NameComplexScript { get; set; }
}
