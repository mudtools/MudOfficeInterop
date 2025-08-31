namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TextFrame 的接口，用于操作文本框格式。
/// </summary>
public interface IWordTextFrame : IDisposable
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
    /// 获取文本框的文本范围。
    /// </summary>
    IWordRange TextRange { get; }

    /// <summary>
    /// 获取或设置文本框的左边距（磅）。
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框的右边距（磅）。
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框的上边距（磅）。
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置文本框的下边距（磅）。
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框的水平对齐方式。
    /// </summary>
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框的垂直对齐方式。
    /// </summary>
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框的自动调整大小方式。
    /// </summary>
    int AutoSize { get; set; }


    /// <summary>
    /// 获取或设置文本框的路径格式。
    /// </summary>
    MsoPathFormat PathFormat { get; set; }

    /// <summary>
    /// 获取文本框的下一文本框。
    /// </summary>
    IWordTextFrame? NextFrame { get; }

    /// <summary>
    /// 获取文本框的上一文本框。
    /// </summary>
    IWordTextFrame? PreviousFrame { get; }

    /// <summary>
    /// 获取文本框的父形状对象。
    /// </summary>
    IWordShape? ParentShape { get; }

    /// <summary>
    /// 获取或设置文本框的文本方向。
    /// </summary>
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取文本框的内部宽度（扣除边距后的宽度）。
    /// </summary>
    float InternalWidth { get; }

    /// <summary>
    /// 获取文本框的内部高度（扣除边距后的高度）。
    /// </summary>
    float InternalHeight { get; }

    /// <summary>
    /// 获取文本框是否包含文本。
    /// </summary>
    bool HasText { get; }

    /// <summary>
    /// 获取文本框中的字符数。
    /// </summary>
    int CharactersCount { get; }

    /// <summary>
    /// 获取文本框中的段落数。
    /// </summary>
    int ParagraphsCount { get; }

    /// <summary>
    /// 获取或设置文本框的填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取或设置文本框的边框格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 连接文本框到下一个文本框。
    /// </summary>
    /// <param name="nextTextFrame">要连接的下一个文本框。</param>
    /// <returns>是否连接成功。</returns>
    bool ConnectTo(IWordTextFrame nextTextFrame);

    /// <summary>
    /// 断开文本框连接。
    /// </summary>
    void BreakLink();

    /// <summary>
    /// 设置文本框边距。
    /// </summary>
    /// <param name="left">左边距。</param>
    /// <param name="right">右边距。</param>
    /// <param name="top">上边距。</param>
    /// <param name="bottom">下边距。</param>
    void SetMargins(float left, float right, float top, float bottom);

    /// <summary>
    /// 设置文本框对齐方式。
    /// </summary>
    /// <param name="horizontal">水平对齐方式。</param>
    /// <param name="vertical">垂直对齐方式。</param>
    void SetAlignment(MsoHorizontalAnchor horizontal, MsoVerticalAnchor vertical);

    /// <summary>
    /// 清除文本框内容。
    /// </summary>
    void ClearText();

    /// <summary>
    /// 复制文本框格式到另一个文本框。
    /// </summary>
    /// <param name="targetTextFrame">目标文本框。</param>
    void CopyTo(IWordTextFrame targetTextFrame);

    /// <summary>
    /// 重置文本框格式为默认值。
    /// </summary>
    void Reset();

    /// <summary>
    /// 获取文本框的文本内容。
    /// </summary>
    /// <returns>文本内容。</returns>
    string GetText();

    /// <summary>
    /// 设置文本框的文本内容。
    /// </summary>
    /// <param name="text">要设置的文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 获取文本框的字体格式。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取文本框的段落格式。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取文本框是否为第一个文本框。
    /// </summary>
    bool IsFirstFrame { get; }

    /// <summary>
    /// 获取文本框是否为最后一个文本框。
    /// </summary>
    bool IsLastFrame { get; }
}