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
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordThreeDFormat : IOfficeObject<IWordThreeDFormat, MsWord.ThreeDFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取或设置形状的挤压深度。可以是 –600 到 9600 之间的值（正值产生前表面为原始形状的挤压；负值产生后表面为原始形状的挤压）。
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// 获取或设置挤压形状围绕 x 轴的旋转角度（以度为单位）。可以是 –90 到 90 之间的值。正值表示向上旋转；负值表示向下旋转。
    /// </summary>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置挤压形状围绕 y 轴的旋转角度（以度为单位）。可以是 –90 到 90 之间的值。正值表示向左旋转；负值表示向右旋转。
    /// </summary>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置挤压形状围绕 z 轴的旋转角度（以度为单位）。
    /// </summary>
    float RotationZ { get; set; }

    /// <summary>
    /// 获取或设置形状在 z 轴上的位置。
    /// </summary>
    float Z { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示挤压颜色是否基于挤压形状的填充（挤压的前表面）并在形状填充更改时自动更改，或者挤压颜色是否独立于形状的填充。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoExtrusionColorType ExtrusionColorType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示挤压是否以透视方式显示（即，挤压的壁是否朝向消失点变窄）。MsoFalse 表示挤压是平行或正交投影（即，挤压的壁不朝向消失点变窄）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置挤压光照强度。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

    /// <summary>
    /// 获取或设置挤压表面材质。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetMaterial PresetMaterial { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定对象或应用于它的格式是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状的轮廓宽度。
    /// </summary>
    float ContourWidth { get; set; }

    /// <summary>
    /// 获取或设置形状的透视量。
    /// </summary>
    float FieldOfView { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示形状上的文本是否随形状旋转。MsoTriState.msoTrue 旋转文本。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ProjectText { get; set; }

    /// <summary>
    /// 获取或设置光照角度。
    /// </summary>
    float LightAngle { get; set; }

    /// <summary>
    /// 获取或设置上斜角类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelTopType { get; set; }

    /// <summary>
    /// 获取或设置上斜角的内嵌大小。
    /// </summary>
    float BevelTopInset { get; set; }

    /// <summary>
    /// 获取或设置上斜角的深度。
    /// </summary>
    float BevelTopDepth { get; set; }

    /// <summary>
    /// 获取或设置下斜角类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBevelType BevelBottomType { get; set; }

    /// <summary>
    /// 获取或设置下斜角的内嵌大小。
    /// </summary>
    float BevelBottomInset { get; set; }

    /// <summary>
    /// 获取或设置下斜角的深度。
    /// </summary>
    float BevelBottomDepth { get; set; }

    /// <summary>
    /// 获取挤压预设方向。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

    /// <summary>
    /// 获取或设置相对于挤压的光源位置。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetLightingDirection PresetLightingDirection { get; set; }

    /// <summary>
    /// 获取或设置光照预设类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLightRigType PresetLighting { get; set; }

    /// <summary>
    /// 获取预设挤压格式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetThreeDFormat PresetThreeDFormat { get; }

    /// <summary>
    /// 获取相机预设类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetCamera PresetCamera { get; }

    /// <summary>
    /// 返回表示形状挤压颜色的 ColorFormat 对象。
    /// </summary>
    IWordColorFormat? ExtrusionColor { get; }

    /// <summary>
    /// 返回表示形状轮廓颜色的 ColorFormat 对象。
    /// </summary>
    IWordColorFormat? ContourColor { get; }

    /// <summary>
    /// 按指定的度数更改指定形状围绕 x 轴的旋转。
    /// </summary>
    /// <param name="increment">必需 Single。指定形状围绕 x 轴的旋转要更改多少（以度为单位）。可以是 –90 到 90 之间的值。正值使形状向上倾斜；负值使其向下倾斜。</param>
    void IncrementRotationX(float increment);

    /// <summary>
    /// 按指定的度数更改指定形状围绕 y 轴的旋转。
    /// </summary>
    /// <param name="increment">必需 Single。指定形状围绕 y 轴的旋转要更改多少（以度为单位）。可以是 –90 到 90 之间的值。正值使形状向左倾斜；负值使其向右倾斜。</param>
    void IncrementRotationY(float increment);

    /// <summary>
    /// 重置挤压围绕 x 轴和 y 轴的旋转为 0（零），使挤压的前表面朝前。此方法不重置围绕 z 轴的旋转。
    /// </summary>
    void ResetRotation();

    /// <summary>
    /// 设置挤压的扫描路径远离挤压形状（挤压的前表面）的方向。
    /// </summary>
    /// <param name="presetExtrusionDirection">必需 MsoPresetExtrusionDirection。可以是以下 MsoPresetExtrusionDirection 常量之一：msoExtrusionTop、msoExtrusionTopRight、msoExtrusionBottom、msoExtrusionBottomLeft、msoExtrusionBottomRight、msoExtrusionLeft、msoExtrusionNone、msoExtrusionRight、msoExtrusionTopLeft、msoPresetExtrusionDirectionMixed。</param>
    void SetExtrusionDirection([ComNamespace("MsCore")] MsoPresetExtrusionDirection presetExtrusionDirection);

    /// <summary>
    /// 设置预设挤压格式。每个预设挤压格式包含一组挤压各种属性的预设值。
    /// </summary>
    /// <param name="presetThreeDFormat">必需 MsoPresetThreeDFormat。指定与单击"绘图"工具栏上的 3-D 按钮时显示的选项（从左到右，从上到下编号）之一对应的预设挤压格式。可以是以下 MsoPresetThreeDFormat 常量之一。注意，为此参数指定 msoPresetThreeDFormatMixed 会导致错误。msoThreeD1、msoThreeD11、msoThreeD13、msoThreeD15、msoThreeD17、msoThreeD19、msoThreeD20、msoThreeD4、msoThreeD6、msoThreeD8、msoPresetThreeDFormatMixed、msoThreeD10、msoThreeD12、msoThreeD14、msoThreeD16、msoThreeD18、msoThreeD2、msoThreeD3、msoThreeD5、msoThreeD7、msoThreeD9。</param>
    void SetThreeDFormat([ComNamespace("MsCore")] MsoPresetThreeDFormat presetThreeDFormat);

    /// <summary>
    /// 设置形状的相机预设。
    /// </summary>
    /// <param name="presetCamera">指定相机预设类型。</param>
    void SetPresetCamera([ComNamespace("MsCore")] MsoPresetCamera presetCamera);

    /// <summary>
    /// 使用指定的增量在 z 轴上旋转形状。
    /// </summary>
    /// <param name="increment">指定增量值。</param>
    void IncrementRotationZ(float increment);

    /// <summary>
    /// 使用指定的增量值在 x 轴上水平旋转形状。
    /// </summary>
    /// <param name="increment">指定增量值。</param>
    void IncrementRotationHorizontal(float increment);

    /// <summary>
    /// 使用指定的增量值在 y 轴上垂直旋转形状。
    /// </summary>
    /// <param name="increment">指定增量值。</param>
    void IncrementRotationVertical(float increment);
}