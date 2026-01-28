//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 序列项接口
/// </summary>
public interface IPowerPointSequence : IDisposable
{
    /// <summary>
    /// 获取效果数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 根据索引获取效果
    /// </summary>
    IPowerPointEffect this[int index] { get; }

    /// <summary>
    /// 添加效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <param name="effectId">效果ID</param>
    /// <param name="trigger">触发器</param>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的效果</returns>
    IPowerPointEffect AddEffect(IPowerPointShape shape, int effectId = 1, int trigger = 1, int index = -1);

    /// <summary>
    /// 删除效果
    /// </summary>
    /// <param name="index">效果索引</param>
    void DeleteEffect(int index);

    /// <summary>
    /// 移动效果
    /// </summary>
    /// <param name="fromIndex">源索引</param>
    /// <param name="toIndex">目标索引</param>
    void MoveEffect(int fromIndex, int toIndex);

    /// <summary>
    /// 查找指定形状的效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <returns>效果列表</returns>
    IEnumerable<IPowerPointEffect> FindEffectsByShape(IPowerPointShape shape);

    /// <summary>
    /// 清除所有效果
    /// </summary>
    void ClearEffects();

    /// <summary>
    /// 设置序列播放时间
    /// </summary>
    /// <param name="startTime">开始时间</param>
    /// <param name="duration">持续时间</param>
    void SetTiming(float startTime, float duration);

    /// <summary>
    /// 获取序列信息
    /// </summary>
    /// <returns>序列信息字符串</returns>
    string GetSequenceInfo();
}
