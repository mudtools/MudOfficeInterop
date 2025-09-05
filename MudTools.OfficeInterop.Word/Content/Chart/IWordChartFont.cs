namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表字体格式的封装接口。
/// </summary>
public interface IWordChartFont : IDisposable
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
    /// 获取或设置字体名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置是否粗体。
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体。
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置是否下划线。
    /// </summary>
    bool Underline { get; set; }

    /// <summary>
    /// 获取或设置字体颜色。
    /// </summary>
    object Color { get; set; }

    /// <summary>
    /// 获取或设置字体颜色索引。
    /// </summary>
    XlColorIndex ColorIndex { get; set; }
}