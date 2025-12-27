//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ThreeDFormat 对象的二次封装接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelThreeDFormat : IOfficeObject<IExcelThreeDFormat?>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置深度
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// 获取或设置 Z 轴坐标值
    /// </summary>
    float Z { get; set; }

    /// <summary>
    /// 获取或设置底部斜角深度
    /// </summary>
    float BevelBottomDepth { get; set; }

    /// <summary>
    /// 获取或设置底部斜角嵌入度
    /// </summary>
    float BevelBottomInset { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角嵌入度
    /// </summary>
    float BevelTopInset { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角深度
    /// </summary>
    float BevelTopDepth { get; set; }

    /// <summary>
    /// 获取或设置挤出颜色类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 获取预设挤出方向
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

    /// <summary>
    /// 获取或设置预设光照方向
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingDirection PresetLightingDirection { get; set; }

    /// <summary>
    /// 获取或设置预设光照柔和度
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

    /// <summary>
    /// 获取或设置预设材料类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetMaterial PresetMaterial { get; set; }

    /// <summary>
    /// 获取预设三维格式类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetThreeDFormat PresetThreeDFormat { get; }

    /// <summary>
    /// 获取或设置预设光照类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLightRigType PresetLighting { get; set; }

    /// <summary>
    /// 获取预设相机类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetCamera PresetCamera { get; }

    /// <summary>
    /// 获取或设置底部斜角类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelBottomType { get; set; }

    /// <summary>
    /// 获取或设置顶部斜角类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelTopType { get; set; }
    /// <summary>
    /// 获取或设置倾斜角度X
    /// </summary>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置倾斜角度Y
    /// </summary>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置倾斜角度Z
    /// </summary>
    float RotationZ { get; set; }

    /// <summary>
    /// 获取或设置轮廓宽度
    /// </summary>
    float ContourWidth { get; set; }

    /// <summary>
    /// 获取或设置视野角度
    /// </summary>
    float FieldOfView { get; set; }

    /// <summary>
    /// 获取或设置光源角度
    /// </summary>
    float LightAngle { get; set; }

    /// <summary>
    /// 获取轮廓颜色格式
    /// </summary>
    IExcelColorFormat? ContourColor { get; }

    /// <summary>
    /// 获取挤出颜色格式
    /// </summary>
    IExcelColorFormat? ExtrusionColor { get; }

    /// <summary>
    /// 获取或设置是否投影文本
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ProjectText { get; set; }

    /// <summary>
    /// 获取或设置透视效果
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置是否可见
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 重置旋转角度为默认值
    /// </summary>
    void ResetRotation();

    /// <summary>
    /// 按指定增量增加 Y 轴旋转角度
    /// </summary>
    /// <param name="increment">要增加的角度值</param>
    void IncrementRotationY(float increment);

    /// <summary>
    /// 按指定增量增加 Z 轴旋转角度
    /// </summary>
    /// <param name="increment">要增加的角度值</param>
    void IncrementRotationZ(float increment);

    /// <summary>
    /// 按指定增量增加水平旋转角度
    /// </summary>
    /// <param name="increment">要增加的角度值</param>
    void IncrementRotationHorizontal(float increment);

    /// <summary>
    /// 按指定增量增加垂直旋转角度
    /// </summary>
    /// <param name="increment">要增加的角度值</param>
    void IncrementRotationVertical(float increment);

    /// <summary>
    /// 设置三维格式预设样式
    /// </summary>
    /// <param name="presetThreeDFormat">预设的三维格式类型</param>
    void SetThreeDFormat([ComNamespace("MsCore")] MsoPresetThreeDFormat presetThreeDFormat);

    /// <summary>
    /// 设置挤出方向预设样式
    /// </summary>
    /// <param name="presetExtrusionDirection">预设的挤出方向类型</param>
    void SetExtrusionDirection([ComNamespace("MsCore")] MsoPresetExtrusionDirection presetExtrusionDirection);

    /// <summary>
    /// 设置预设相机视角
    /// </summary>
    /// <param name="presetCamera">预设的相机视角类型</param>
    void SetPresetCamera([ComNamespace("MsCore")] MsoPresetCamera presetCamera);
}