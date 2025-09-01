//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shading 的接口，用于操作段落或表格的底纹样式。
/// </summary>
public interface IWordShading : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置底纹的背景颜色（RGB值）。
    /// </summary>
    WdColor BackgroundPatternColor { get; set; }

    /// <summary>
    /// 获取或设置底纹的前景颜色（RGB值）。
    /// </summary>
    WdColor ForegroundPatternColor { get; set; }

    /// <summary>
    /// 获取或设置底纹图案样式。
    /// </summary>
    WdTextureIndex Texture { get; set; }

    /// <summary>
    /// 获取或设置底纹的颜色索引。
    /// </summary>
    WdColorIndex BackgroundPatternColorIndex { get; set; }

    /// <summary>
    /// 获取或设置前景颜色索引。
    /// </summary>
    WdColorIndex ForegroundPatternColorIndex { get; set; }

    /// <summary>
    /// 清除底纹设置。
    /// </summary>
    void Clear();

    /// <summary>
    /// 应用纯色底纹。
    /// </summary>
    /// <param name="color">背景颜色（RGB值）。</param>
    void ApplySolidColor(WdColor color);

    /// <summary>
    /// 应用纹理底纹。
    /// </summary>
    /// <param name="texture">纹理样式。</param>
    void ApplyTexture(WdTextureIndex texture);

    /// <summary>
    /// 复制底纹设置到另一个对象。
    /// </summary>
    /// <param name="targetShading">目标底纹对象。</param>
    void CopyTo(IWordShading targetShading);
}