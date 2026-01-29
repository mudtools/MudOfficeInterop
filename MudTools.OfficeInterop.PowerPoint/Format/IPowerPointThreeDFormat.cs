//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 形状的三维格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointThreeDFormat : IDisposable
{
    /// <summary>
    /// 获取创建此三维格式设置的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此三维格式设置的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此三维格式设置的父对象。
    /// </summary>
    /// <value>表示此三维格式设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 按指定增量增加绕 X 轴的旋转角度。
    /// </summary>
    /// <param name="increment">旋转角度增量（以度为单位）。</param>
    void IncrementRotationX(float increment);

    /// <summary>
    /// 按指定增量增加绕 Y 轴的旋转角度。
    /// </summary>
    /// <param name="increment">旋转角度增量（以度为单位）。</param>
    void IncrementRotationY(float increment);

    /// <summary>
    /// 重置所有旋转角度为默认值。
    /// </summary>
    void ResetRotation();

    /// <summary>
    /// 应用预定义的三维格式。
    /// </summary>
    /// <param name="presetThreeDFormat">要应用的三维格式预设。</param>
    void SetThreeDFormat([ComNamespace("MsCore")] MsoPresetThreeDFormat presetThreeDFormat);

    /// <summary>
    /// 设置拉伸方向。
    /// </summary>
    /// <param name="presetExtrusionDirection">要应用的拉伸方向预设。</param>
    void SetExtrusionDirection([ComNamespace("MsCore")] MsoPresetExtrusionDirection presetExtrusionDirection);

    /// <summary>
    /// 获取或设置三维形状的深度（以磅为单位）。
    /// </summary>
    /// <value>表示深度的浮点数。</value>
    float Depth { get; set; }

    /// <summary>
    /// 获取拉伸部分的颜色设置。
    /// </summary>
    /// <value>表示拉伸颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? ExtrusionColor { get; }

    /// <summary>
    /// 获取或设置拉伸部分的颜色类型。
    /// </summary>
    /// <value>表示拉伸颜色类型的 <see cref="MsoExtrusionColorType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用透视效果。
    /// </summary>
    /// <value>指示是否启用透视效果的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Perspective { get; set; }

    /// <summary>
    /// 获取当前拉伸方向的预设类型。
    /// </summary>
    /// <value>表示拉伸方向预设的 <see cref="MsoPresetExtrusionDirection"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

    /// <summary>
    /// 获取或设置光源方向的预设类型。
    /// </summary>
    /// <value>表示光源方向预设的 <see cref="MsoPresetLightingDirection"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingDirection PresetLightingDirection { get; set; }

    /// <summary>
    /// 获取或设置光源柔和度的预设类型。
    /// </summary>
    /// <value>表示光源柔和度预设的 <see cref="MsoPresetLightingSoftness"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

    /// <summary>
    /// 获取或设置材质表面的预设类型。
    /// </summary>
    /// <value>表示材质预设的 <see cref="MsoPresetMaterial"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetMaterial PresetMaterial { get; set; }

    /// <summary>
    /// 获取当前三维格式的预设类型。
    /// </summary>
    /// <value>表示三维格式预设的 <see cref="MsoPresetThreeDFormat"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetThreeDFormat PresetThreeDFormat { get; }

    /// <summary>
    /// 获取或设置绕 X 轴的旋转角度。
    /// </summary>
    /// <value>表示 X 轴旋转角度的浮点数。</value>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置绕 Y 轴的旋转角度。
    /// </summary>
    /// <value>表示 Y 轴旋转角度的浮点数。</value>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示三维效果是否可见。
    /// </summary>
    /// <value>指示三维效果是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 设置摄像机的预设类型。
    /// </summary>
    /// <param name="presetCamera">要应用的摄像机预设。</param>
    void SetPresetCamera([ComNamespace("MsCore")] MsoPresetCamera presetCamera);

    /// <summary>
    /// 按指定增量增加绕 Z 轴的旋转角度。
    /// </summary>
    /// <param name="increment">旋转角度增量（以度为单位）。</param>
    void IncrementRotationZ(float increment);

    /// <summary>
    /// 按指定增量增加水平旋转角度。
    /// </summary>
    /// <param name="increment">旋转角度增量（以度为单位）。</param>
    void IncrementRotationHorizontal(float increment);

    /// <summary>
    /// 按指定增量增加垂直旋转角度。
    /// </summary>
    /// <param name="increment">旋转角度增量（以度为单位）。</param>
    void IncrementRotationVertical(float increment);

    /// <summary>
    /// 获取或设置光源组合的预设类型。
    /// </summary>
    /// <value>表示光源组合预设的 <see cref="MsoLightRigType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLightRigType PresetLighting { get; set; }

    /// <summary>
    /// 获取或设置形状在 Z 轴上的位置。
    /// </summary>
    /// <value>表示 Z 轴位置的浮点数。</value>
    float Z { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角类型。
    /// </summary>
    /// <value>表示顶部斜角类型的 <see cref="MsoBevelType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelTopType { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角的内切量（以磅为单位）。
    /// </summary>
    /// <value>表示顶部斜角内切量的浮点数。</value>
    float BevelTopInset { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角的深度（以磅为单位）。
    /// </summary>
    /// <value>表示顶部斜角深度的浮点数。</value>
    float BevelTopDepth { get; set; }

    /// <summary>
    /// 获取或设置底部斜角类型。
    /// </summary>
    /// <value>表示底部斜角类型的 <see cref="MsoBevelType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelBottomType { get; set; }

    /// <summary>
    /// 获取或设置底部斜角的内切量（以磅为单位）。
    /// </summary>
    /// <value>表示底部斜角内切量的浮点数。</value>
    float BevelBottomInset { get; set; }

    /// <summary>
    /// 获取或设置底部斜角的深度（以磅为单位）。
    /// </summary>
    /// <value>表示底部斜角深度的浮点数。</value>
    float BevelBottomDepth { get; set; }

    /// <summary>
    /// 获取当前摄像机的预设类型。
    /// </summary>
    /// <value>表示摄像机预设的 <see cref="MsoPresetCamera"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetCamera PresetCamera { get; }

    /// <summary>
    /// 获取或设置绕 Z 轴的旋转角度。
    /// </summary>
    /// <value>表示 Z 轴旋转角度的浮点数。</value>
    float RotationZ { get; set; }

    /// <summary>
    /// 获取或设置轮廓线的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示轮廓线宽度的浮点数。</value>
    float ContourWidth { get; set; }

    /// <summary>
    /// 获取轮廓线的颜色设置。
    /// </summary>
    /// <value>表示轮廓线颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? ContourColor { get; }

    /// <summary>
    /// 获取或设置摄像机的视野角度。
    /// </summary>
    /// <value>表示视野角度的浮点数。</value>
    float FieldOfView { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否投影文本到三维表面。
    /// </summary>
    /// <value>指示是否投影文本的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ProjectText { get; set; }

    /// <summary>
    /// 获取或设置光源的角度。
    /// </summary>
    /// <value>表示光源角度的浮点数。</value>
    float LightAngle { get; set; }
}