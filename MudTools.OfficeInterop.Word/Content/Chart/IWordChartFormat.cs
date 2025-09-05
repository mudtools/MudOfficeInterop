namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表格式的封装接口。
/// </summary>
public interface IWordChartFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取线条格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取阴影格式。
    /// </summary>
    IWordShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取发光格式。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取三维格式。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }
}