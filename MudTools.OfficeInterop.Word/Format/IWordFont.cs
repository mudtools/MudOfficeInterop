namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Font 的接口，用于操作 Word 文档中文字的字体样式。
/// </summary>
public interface IWordFont : IDisposable
{
    /// <summary>
    /// 获取或设置字体名称（如 "Arial", "Times New Roman" 等）。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小（单位为磅）。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置是否为粗体。
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否为斜体。
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置是否带下划线。
    /// </summary>
    bool Underline { get; set; }

    /// <summary>
    /// 获取或设置字体颜色（RGB 值）。
    /// </summary>
    WdColor Color { get; set; }

    /// <summary>
    /// 获取或设置上标（如数学中的平方）。
    /// </summary>
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置下标（如化学式 H₂O）。
    /// </summary>
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置字符间距（单位为磅）。
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置字符缩放比例（百分比，如 100 表示正常大小）。
    /// </summary>
    int Scaling { get; set; }

    /// <summary>
    /// 获取或设置字符位置偏移（正值为上移，负值为下移）。
    /// </summary>
    int Position { get; set; }
}