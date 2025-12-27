//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 动画效果接口
/// </summary>
public interface IPowerPointEffect : IDisposable
{
    /// <summary>
    /// 获取目标形状
    /// </summary>
    IPowerPointShape Shape { get; }

    /// <summary>
    /// 获取或设置效果类型
    /// </summary>
    int EffectType { get; set; }

    /// <summary>
    /// 获取效果信息
    /// </summary>
    IPowerPointEffectInformation EffectInformation { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置效果索引
    /// </summary>
    int Index { get; set; }


    /// <summary>
    /// 应用效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="triggerType">触发类型</param>
    void ApplyEffect(int effectType, int triggerType = 1);

    /// <summary>
    /// 删除效果
    /// </summary>
    void Delete();

    /// <summary>
    /// 移动效果到指定位置
    /// </summary>
    /// <param name="index">目标位置</param>
    void MoveTo(int index);

    /// <summary>
    /// 设置效果参数
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <param name="value">属性值</param>
    void SetProperty(string propertyName, object value);

    /// <summary>
    /// 获取效果参数
    /// </summary>
    /// <param name="propertyName">属性名称</param>
    /// <returns>属性值</returns>
    object GetProperty(string propertyName);

    /// <summary>
    /// 预览效果
    /// </summary>
    void Preview();

    /// <summary>
    /// 获取效果信息
    /// </summary>
    /// <returns>效果信息字符串</returns>
    string GetEffectInfo();
}