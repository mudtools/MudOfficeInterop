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