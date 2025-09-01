//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

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