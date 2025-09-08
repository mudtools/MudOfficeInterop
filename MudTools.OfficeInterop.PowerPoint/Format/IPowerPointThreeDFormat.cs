//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 三维格式接口
/// </summary>
public interface IPowerPointThreeDFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置深度
    /// </summary>
    float Depth { get; set; }

    /// <summary>
    /// 获取或设置挤出颜色
    /// </summary>
    int ExtrusionColor { get; set; }

    /// <summary>
    /// 获取或设置预设光照
    /// </summary>
    int PresetLighting { get; set; }

    /// <summary>
    /// 获取或设置预设材质
    /// </summary>
    int PresetMaterial { get; set; }

    /// <summary>
    /// 获取或设置预设三维格式
    /// </summary>
    int PresetThreeDFormat { get; set; }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置透视效果
    /// </summary>
    bool Perspective { get; set; }

    /// <summary>
    /// 获取或设置旋转X轴角度
    /// </summary>
    float RotationX { get; set; }

    /// <summary>
    /// 获取或设置旋转Y轴角度
    /// </summary>
    float RotationY { get; set; }

    /// <summary>
    /// 获取或设置旋转Z轴角度
    /// </summary>
    float RotationZ { get; set; }

    /// <summary>
    /// 获取或设置光照角度
    /// </summary>
    float LightAngle { get; set; }

    /// <summary>
    /// 设置预设三维格式
    /// </summary>
    /// <param name="presetThreeDFormat">预设三维格式</param>
    void SetThreeDFormat(int presetThreeDFormat);

    /// <summary>
    /// 设置挤出方向
    /// </summary>
    /// <param name="presetExtrusionDirection">预设挤出方向</param>
    void SetExtrusionDirection(int presetExtrusionDirection);

    /// <summary>
    /// 设置光照效果
    /// </summary>
    /// <param name="presetLighting">预设光照</param>
    /// <param name="lightAngle">光照角度</param>
    void SetLighting(int presetLighting, float lightAngle = 0);

    /// <summary>
    /// 设置材质效果
    /// </summary>
    /// <param name="presetMaterial">预设材质</param>
    void SetMaterial(int presetMaterial);

    /// <summary>
    /// 设置旋转角度
    /// </summary>
    /// <param name="rotationX">X轴旋转角度</param>
    /// <param name="rotationY">Y轴旋转角度</param>
    /// <param name="rotationZ">Z轴旋转角度</param>
    void SetRotation(float rotationX = 0, float rotationY = 0, float rotationZ = 0);

    /// <summary>
    /// 重置三维格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制三维格式
    /// </summary>
    /// <returns>复制的三维格式对象</returns>
    IPowerPointThreeDFormat Duplicate();

    /// <summary>
    /// 应用三维格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    void ApplyTo(IPowerPointShape shape);

    /// <summary>
    /// 设置深度和挤出颜色
    /// </summary>
    /// <param name="depth">深度</param>
    /// <param name="extrusionColor">挤出颜色</param>
    void SetDepthAndColor(float depth, int extrusionColor);

    /// <summary>
    /// 设置透视效果
    /// </summary>
    /// <param name="perspective">是否启用透视</param>
    /// <param name="autoRotation">是否自动旋转</param>
    void SetPerspective(bool perspective, bool autoRotation = false);

    /// <summary>
    /// 获取三维信息
    /// </summary>
    /// <returns>三维信息字符串</returns>
    string GetThreeDInfo();
}
