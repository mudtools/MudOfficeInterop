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
public interface IOfficeThreeDFormat : IDisposable
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
    /// 获取或设置三维效果的透视效果启用状态。
    /// </summary>
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置三维效果的可见性。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置三维效果的Z轴坐标值。
    /// </summary>
    float Z { get; set; }

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
    IOfficeColorFormat ExtrusionColor { get; }

    /// <summary>
    /// 获取或设置三维效果的轮廓颜色类型。
    /// </summary>
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 设置三维效果的透视相机。
    /// </summary>
    /// <param name="presetCamera">预设相机类型。</param>
    void SetThreeDFormat(MsoPresetThreeDFormat presetCamera);

    /// <summary>
    /// 应用预设的三维效果。
    /// </summary>
    /// <param name="presetThreeDFormat">预设三维效果。</param>
    void PresetThreeDFormat(MsoPresetExtrusionDirection presetThreeDFormat);

    /// <summary>
    /// 设置灯光效果。
    /// </summary>
    void SetLightRig(MsoPresetCamera presetCamera);

    /// <summary>
    /// 重置三维格式。
    /// </summary>
    void ResetRotation();
}