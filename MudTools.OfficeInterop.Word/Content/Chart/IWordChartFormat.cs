namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表格式的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartFormat : IOfficeObject<IWordChartFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
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

    /// <summary>
    /// 获取图片格式。
    /// </summary>
    IWordPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取柔边格式。
    /// </summary>
    IWordSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取文本框格式。
    /// </summary>
    IOfficeTextFrame2? TextFrame2 { get; }
}