//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.ThreeDFormat 的接口，用于操作三维格式。
/// </summary>
public interface IWordThreeDFormat : IDisposable
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
    /// 获取或设置轮廓线宽度。
    /// </summary>
    float ContourWidth { get; set; }

    /// <summary>
    /// 获取或设置顶部斜面类型。
    /// </summary>
    MsoBevelType BevelTopType { get; set; }

    /// <summary>
    /// 获取或设置顶部斜面深度。
    /// </summary>
    float BevelTopDepth { get; set; }

    /// <summary>
    /// 获取或设置顶部斜面插入宽度。
    /// </summary>
    float BevelTopInset { get; set; }

    /// <summary>
    /// 获取或设置底部斜面深度。
    /// </summary>
    float BevelBottomDepth { get; set; }

    /// <summary>
    /// 获取或设置底部斜面插入宽度。
    /// </summary>
    float BevelBottomInset { get; set; }

    /// <summary>
    /// 获取或设置底部斜面类型。
    /// </summary>
    MsoBevelType BevelBottomType { get; set; }

    /// <summary>
    /// 获取或设置 Z 轴坐标值。
    /// </summary>
    float Z { get; set; }

    /// <summary>
    /// 获取轮廓线颜色格式。
    /// </summary>
    IWordColorFormat? ContourColor { get; }

    /// <summary>
    /// 获取或设置视野角度。
    /// </summary>
    float FieldOfView { get; set; }

    /// <summary>
    /// 获取或设置光源角度。
    /// </summary>
    float LightAngle { get; set; }

    /// <summary>
    /// 获取或设置三维形状的深度（磅）。
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// 获取或设置三维形状的 extrusion 颜色格式。
    /// </summary>
    IWordColorFormat? ExtrusionColor { get; }

    /// <summary>
    /// 获取或设置 extrusion 颜色类型。
    /// </summary>
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 获取或设置透视效果。
    /// </summary>
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置三维效果是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置是否将文本投影到三维形状上。
    /// </summary>
    bool ProjectText { get; set; }


    /// <summary>
    /// 获取预设相机效果。
    /// </summary>
    MsoPresetCamera PresetCamera { get; }

    /// <summary>
    /// 获取或设置预设光照柔和度。
    /// </summary>
    MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

    /// <summary>
    /// 获取或设置预设光照方向。
    /// </summary>
    MsoPresetLightingDirection PresetLightingDirection { get; set; }

    /// <summary>
    /// 获取或设置材质预设类型。
    /// </summary>
    MsoPresetMaterial PresetMaterial { get; set; }

    /// <summary>
    /// 获取或设置光照强度。
    /// </summary>
    MsoLightRigType PresetLighting { get; set; }

    /// <summary>
    /// 获取或设置水平旋转角度。
    /// </summary>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置垂直旋转角度。
    /// </summary>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置旋转角度（Z轴）。
    /// </summary>
    float RotationZ { get; set; }

    /// <summary>
    /// 获取或设置是否应用透视效果。
    /// </summary>
    bool PerspectiveEnabled { get; set; }

    /// <summary>
    /// 设置预设三维效果。
    /// </summary>
    /// <param name="presetThreeDFormat">预设三维格式类型。</param>
    void SetPresetCamera(MsoPresetCamera presetThreeDFormat);

    /// <summary>
    /// 设置预设光照效果。
    /// </summary>
    /// <param name="presetLighting">预设光照类型。</param>
    void SetPresetLighting(MsoLightRigType presetLighting);

    /// <summary>
    /// 设置预设材质效果。
    /// </summary>
    /// <param name="presetMaterial">预设材质类型。</param>
    void SetPresetMaterial(MsoPresetMaterial presetMaterial);

    /// <summary>
    /// 设置预设 extrusion 方向。
    /// </summary>
    /// <param name="presetExtrusionDirection">预设 extrusion 方向。</param>
    void SetExtrusionDirection(MsoPresetExtrusionDirection presetExtrusionDirection);

    /// <summary>
    /// 设置旋转角度。
    /// </summary>
    /// <param name="rotationX">X轴旋转角度。</param>
    /// <param name="rotationY">Y轴旋转角度。</param>
    /// <param name="rotationZ">Z轴旋转角度。</param>
    void SetRotation(float rotationX, float rotationY, float rotationZ);

    /// <summary>
    /// 清除三维格式效果。
    /// </summary>
    void Clear();

    /// <summary>
    /// 复制三维格式到另一个对象。
    /// </summary>
    /// <param name="targetThreeD">目标三维格式对象。</param>
    void CopyTo(IWordThreeDFormat targetThreeD);

    /// <summary>
    /// 设置 extrusion 颜色。
    /// </summary>
    /// <param name="color">RGB颜色值。</param>
    void SetExtrusionColor(int color);
}