//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中三维格式的接口封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeThreeDFormat : IOfficeObject<IOfficeThreeDFormat>, IDisposable
{
    /// <summary>
    /// 获取或设置三维效果的深度。
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// 获取或设置三维效果的倾斜角度。
    /// </summary>
    float BevelTopInset { get; set; }

    /// <summary>
    /// 获取或设置顶部斜面的高度。
    /// </summary>
    float BevelTopDepth { get; set; }

    /// <summary>
    /// 获取或设置底部斜面的插入深度。
    /// </summary>
    float BevelBottomInset { get; set; }

    /// <summary>
    /// 获取或设置底部斜面的深度。
    /// </summary>
    float BevelBottomDepth { get; set; }

    /// <summary>
    /// 获取或设置底部斜面类型。
    /// </summary>
    MsoBevelType BevelBottomType { get; set; }

    /// <summary>
    /// 获取或设置顶部斜面类型。
    /// </summary>
    MsoBevelType BevelTopType { get; set; }

    /// <summary>
    /// 获取或设置预设光源类型。
    /// </summary>
    MsoLightRigType PresetLighting { get; set; }

    /// <summary>
    /// 获取预设的三维格式类型。
    /// </summary>
    MsoPresetThreeDFormat PresetThreeDFormat { get; }

    /// <summary>
    /// 获取预设的挤出方向。
    /// </summary>
    MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

    /// <summary>
    /// 获取或设置三维效果的透视效果启用状态。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置三维效果的可见性。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置三维效果的Z轴坐标值。
    /// </summary>
    float Z { get; set; }


    /// <summary>
    /// 获取或设置三维效果的旋转X轴角度。
    /// </summary>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置三维效果的旋转Y轴角度。
    /// </summary>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置三维效果的旋转Z轴角度。
    /// </summary>
    float RotationZ { get; set; }


    /// <summary>
    /// 获取或设置三维效果的场深。
    /// </summary>
    float FieldOfView { get; set; }

    /// <summary>
    /// 获取或设置三维效果的灯光角度。
    /// </summary>
    float LightAngle { get; set; }

    /// <summary>
    /// 获取或设置三维效果的材质类型。
    /// </summary>
    MsoPresetMaterial PresetMaterial { get; set; }

    /// <summary>
    /// 获取或设置三维效果的光照效果。
    /// </summary>
    MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

    /// <summary>
    /// 获取或设置三维效果的光照方向。
    /// </summary>
    MsoPresetLightingDirection PresetLightingDirection { get; set; }

    /// <summary>
    /// 获取或设置三维效果的轮廓颜色。
    /// </summary>
    IOfficeColorFormat? ExtrusionColor { get; }

    /// <summary>
    /// 获取或设置三维效果的轮廓颜色类型。
    /// </summary>
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 获取三维形状的轮廓颜色。
    /// </summary>
    IOfficeColorFormat? ContourColor { get; }

    /// <summary>
    /// 获取或设置三维形状的轮廓宽度。
    /// </summary>
    float ContourWidth { get; set; }

    /// <summary>
    /// 获取三维形状的预设相机效果。
    /// </summary>
    MsoPresetCamera PresetCamera { get; }

    /// <summary>
    /// 水平旋转角度增量调整。
    /// </summary>
    /// <param name="increment">水平旋转角度的增量值</param>
    void IncrementRotationHorizontal(float increment);

    /// <summary>
    /// X轴旋转角度增量调整。
    /// </summary>
    /// <param name="increment">X轴旋转角度的增量值</param>
    void IncrementRotationX(float increment);

    /// <summary>
    /// Y轴旋转角度增量调整。
    /// </summary>
    /// <param name="increment">Y轴旋转角度的增量值</param>
    void IncrementRotationY(float increment);

    /// <summary>
    /// 设置三维效果的透视相机。
    /// </summary>
    /// <param name="presetCamera">预设相机类型。</param>
    void SetThreeDFormat(MsoPresetThreeDFormat presetCamera);

    /// <summary>
    /// 设置预设相机效果。
    /// </summary>
    /// <param name="PresetCamera">预设相机类型。</param>
    void SetPresetCamera(MsoPresetCamera PresetCamera);

    /// <summary>
    /// Z轴旋转角度增量调整。
    /// </summary>
    /// <param name="increment">Z轴旋转角度的增量值</param>
    void IncrementRotationZ(float increment);


    /// <summary>
    /// 重置三维格式。
    /// </summary>
    void ResetRotation();
}