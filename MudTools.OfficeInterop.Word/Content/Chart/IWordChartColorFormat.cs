namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表颜色格式的封装接口。
/// </summary>
public interface IWordChartColorFormat : IDisposable
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
    /// 获取或设置 RGB 颜色值。
    /// </summary>
    int RGB { get; }

    /// <summary>
    /// 获取颜色类型。
    /// </summary>
    int Type { get; }
}