//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Frame 的接口，用于操作框架格式。
/// </summary>
public interface IWordFrame : IDisposable
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
    /// 获取框架所在的范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取框架的底纹格式设置。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置框架的水平位置。
    /// </summary>
    float HorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置框架的垂直位置。
    /// </summary>
    float VerticalPosition { get; set; }

    /// <summary>
    /// 获取或设置框架的水平距离（磅）。
    /// </summary>
    float HorizontalDistanceFromText { get; set; }

    /// <summary>
    /// 获取或设置框架的垂直距离（磅）。
    /// </summary>
    float VerticalDistanceFromText { get; set; }

    /// <summary>
    /// 获取或设置框架的宽度（磅）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置框架的高度（磅）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置框架是否锁定锚点。
    /// </summary>
    bool LockAnchor { get; set; }

    /// <summary>
    /// 获取或设置框架的文本环绕方式。
    /// </summary>
    bool TextWrap { get; set; }

    /// <summary>
    /// 获取框架是否包含文本。
    /// </summary>
    bool HasText { get; }

    /// <summary>
    /// 获取框架中的字符数。
    /// </summary>
    int CharactersCount { get; }

    /// <summary>
    /// 获取框架中的段落数。
    /// </summary>
    int ParagraphsCount { get; }

    /// <summary>
    /// 删除框架。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择框架。
    /// </summary>
    void Select();

    /// <summary>
    /// 设置框架与文本的距离。
    /// </summary>
    /// <param name="horizontal">水平距离。</param>
    /// <param name="vertical">垂直距离。</param>
    void SetDistanceFromText(float horizontal, float vertical);

    /// <summary>
    /// 调整框架大小。
    /// </summary>
    /// <param name="width">新宽度。</param>
    /// <param name="height">新高度。</param>
    void Resize(float width, float height);

    /// <summary>
    /// 移动框架到指定的Z轴位置。
    /// </summary>
    /// <param name="position">Z轴位置。</param>
    void ZOrder(MsoZOrderCmd position);

    /// <summary>
    /// 连接框架到下一个框架。
    /// </summary>
    /// <param name="nextFrame">要连接的下一个框架。</param>
    /// <returns>是否连接成功。</returns>
    bool ConnectTo(IWordFrame nextFrame);

    /// <summary>
    /// 断开框架连接。
    /// </summary>
    void BreakLink();

    /// <summary>
    /// 复制框架格式到另一个框架。
    /// </summary>
    /// <param name="targetFrame">目标框架。</param>
    void CopyTo(IWordFrame targetFrame);

    /// <summary>
    /// 重置框架格式为默认值。
    /// </summary>
    void Reset();

    /// <summary>
    /// 获取框架的文本内容。
    /// </summary>
    /// <returns>文本内容。</returns>
    string GetText();

    /// <summary>
    /// 设置框架的文本内容。
    /// </summary>
    /// <param name="text">要设置的文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 获取框架的字体格式。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取框架的段落格式。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }
}