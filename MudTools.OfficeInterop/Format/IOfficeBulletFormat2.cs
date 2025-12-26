//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 应用程序中文本项目符号格式的接口封装。
/// 该接口提供了对项目符号的各种属性和行为的访问，包括样式、类型、字体、可见性等设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeBulletFormat2 : IOfficeObject<IOfficeBulletFormat2>, IDisposable
{
    /// <summary>
    /// 获取项目符号格式的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置项目符号相对于文本的大小比例。
    /// </summary>
    float RelativeSize { get; set; }

    /// <summary>
    /// 获取或设置项目符号的编号样式。
    /// </summary>
    MsoNumberedBulletStyle Style { get; set; }

    /// <summary>
    /// 获取或设置项目符号的类型（无符号、未编号、编号或图片）。
    /// </summary>
    MsoBulletType Type { get; set; }

    /// <summary>
    /// 获取或设置是否使用文本颜色作为项目符号颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseTextColor { get; set; }

    /// <summary>
    /// 获取或设置是否使用文本字体作为项目符号字体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseTextFont { get; set; }

    /// <summary>
    /// 获取或设置项目符号是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置项目符号编号的起始值。
    /// </summary>
    int StartValue { get; set; }

    /// <summary>
    /// 获取或设置用作项目符号的字符的 Unicode 编码值。
    /// </summary>
    int Character { get; set; }

    /// <summary>
    /// 获取项目符号的字体格式设置。
    /// </summary>
    IOfficeFont2? Font { get; }

    /// <summary>
    /// 获取项目符号的编号值（仅适用于编号项目符号）。
    /// </summary>
    int Number { get; }

    /// <summary>
    /// 设置用作项目符号的图片。
    /// </summary>
    /// <param name="fileName">图片文件的完整路径。</param>
    void Picture(string fileName);
}