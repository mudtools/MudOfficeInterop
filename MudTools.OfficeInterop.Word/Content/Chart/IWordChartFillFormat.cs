//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表填充格式的封装接口。
/// </summary>
public interface IWordChartFillFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置填充前景色。
    /// </summary>
    IWordChartColorFormat ForeColor { get; }

    /// <summary>
    /// 获取或设置填充背景色。
    /// </summary>
    IWordChartColorFormat BackColor { get; }

    /// <summary>
    /// 获取或设置填充样式。
    /// </summary>
    MsoPatternType Pattern { get; }

    /// <summary>
    /// 获取或设置填充类型。
    /// </summary>
    MsoFillType Type { get; }

    /// <summary>
    /// 设置渐变填充。
    /// </summary>
    /// <param name="style">渐变样式。</param>
    /// <param name="variant">渐变变体。</param>
    /// <param name="degree">渐变角度。</param>
    void OneColorGradient(MsoGradientStyle style, int variant, float degree);

    /// <summary>
    /// 设置纹理填充。
    /// </summary>
    /// <param name="texture">纹理类型。</param>
    void PresetTextured(MsoPresetTexture texture);
}