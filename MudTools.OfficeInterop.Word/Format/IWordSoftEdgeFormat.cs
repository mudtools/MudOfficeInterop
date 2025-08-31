namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.SoftEdgeFormat 的接口，用于操作柔化边缘格式。
/// </summary>
public interface IWordSoftEdgeFormat : IDisposable
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
    /// 获取或设置柔化边缘的类型。
    /// </summary>
    MsoSoftEdgeType Type { get; set; }

    /// <summary>
    /// 获取或设置柔化边缘的半径（磅）。
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// 获取柔化边缘是否可见。
    /// </summary>
    bool Visible { get; }

    /// <summary>
    /// 获取柔化边缘的大小。
    /// </summary>
    float Size { get; }

    /// <summary>
    /// 应用预设柔化边缘效果。
    /// </summary>
    /// <param name="softEdgeType">预设柔化边缘类型。</param>
    void ApplyPreset(MsoSoftEdgeType softEdgeType);

    /// <summary>
    /// 设置自定义柔化边缘效果。
    /// </summary>
    /// <param name="radius">柔化半径。</param>
    /// <param name="transparency">透明度。</param>
    void SetCustomSoftEdge(float radius, float transparency = 0.5f);

    /// <summary>
    /// 清除柔化边缘效果。
    /// </summary>
    void Clear();

    /// <summary>
    /// 复制柔化边缘格式到另一个对象。
    /// </summary>
    /// <param name="targetSoftEdge">目标柔化边缘格式对象。</param>
    void CopyTo(IWordSoftEdgeFormat targetSoftEdge);

    /// <summary>
    /// 重置柔化边缘格式为默认值。
    /// </summary>
    void Reset();

    /// <summary>
    /// 验证柔化边缘参数是否有效。
    /// </summary>
    /// <param name="radius">柔化半径。</param>
    /// <param name="transparency">透明度。</param>
    /// <returns>参数是否有效。</returns>
    bool ValidateParameters(float radius, float transparency);

    /// <summary>
    /// 获取是否应用了柔化边缘效果。
    /// </summary>
    bool HasSoftEdgeEffect { get; }
}