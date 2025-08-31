namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ConditionalStyle 的接口，用于操作条件样式。
/// 条件样式用于定义表格中特定部分（如标题行、奇数行等）的格式。
/// </summary>
public interface IWordConditionalStyle : IDisposable
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
    /// 获取条件样式的边框集合。
    /// </summary>
    IWordBorders Borders { get; }

    /// <summary>
    /// 获取条件样式的底纹。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 获取条件样式的字体格式。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取条件样式的段落格式。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }

    float BottomPadding { get; set; }

    float TopPadding { get; set; }

    float LeftPadding { get; set; }

    float RightPadding { get; set; }
}