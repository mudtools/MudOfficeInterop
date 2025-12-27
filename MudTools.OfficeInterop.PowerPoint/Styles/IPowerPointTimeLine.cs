//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 时间线接口
/// </summary>
public interface IPowerPointTimeLine : IDisposable
{
    /// <summary>
    /// 获取动画序列集合
    /// </summary>
    IPowerPointSequences Sequences { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置是否启用动画
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取动画效果数量
    /// </summary>
    int EffectCount { get; }

    /// <summary>
    /// 获取主序列
    /// </summary>
    IPowerPointSequence MainSequence { get; }

    /// <summary>
    /// 获取交互序列
    /// </summary>
    IPowerPointSequence InteractiveSequences { get; }

    /// <summary>
    /// 添加动画序列
    /// </summary>
    /// <param name="index">插入位置</param>
    /// <returns>新添加的序列</returns>
    IPowerPointSequence AddSequence(int index = -1);

    /// <summary>
    /// 刷新动画显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 应用动画方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    void ApplyAnimationScheme(int schemeIndex = -1);

    /// <summary>
    /// 复制动画到其他幻灯片
    /// </summary>
    /// <param name="targetSlide">目标幻灯片</param>
    void CopyTo(IPowerPointSlide targetSlide);

    /// <summary>
    /// 获取动画效果
    /// </summary>
    /// <param name="index">效果索引</param>
    /// <returns>动画效果</returns>
    IPowerPointEffect GetEffect(int index);

    /// <summary>
    /// 查找指定形状的动画效果
    /// </summary>
    /// <param name="shape">目标形状</param>
    /// <returns>动画效果列表</returns>
    IEnumerable<IPowerPointEffect> FindEffectsByShape(IPowerPointShape shape);

    /// <summary>
    /// 设置动画播放顺序
    /// </summary>
    /// <param name="effectOrder">效果顺序数组</param>
    void SetEffectOrder(int[] effectOrder);

    /// <summary>
    /// 获取时间线信息
    /// </summary>
    /// <returns>时间线信息字符串</returns>
    string GetTimeLineInfo();
}
