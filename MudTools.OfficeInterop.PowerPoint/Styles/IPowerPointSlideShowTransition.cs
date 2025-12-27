//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 幻灯片放映切换效果接口
/// </summary>
public interface IPowerPointSlideShowTransition : IDisposable
{
    /// <summary>
    /// 获取或设置进入效果
    /// </summary>
    int EntryEffect { get; set; }

    /// <summary>
    /// 获取或设置是否定时前进
    /// </summary>
    int AdvanceOnTime { get; set; }

    /// <summary>
    /// 获取或设置前进时间
    /// </summary>
    float AdvanceTime { get; set; }

    /// <summary>
    /// 获取或设置是否隐藏幻灯片
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置声音效果
    /// </summary>
    string SoundEffect { get; set; }

    /// <summary>
    /// 获取或设置持续时间
    /// </summary>
    float Duration { get; set; }

    /// <summary>
    /// 获取或设置速度
    /// </summary>
    int Speed { get; set; }


    /// <summary>
    /// 获取或设置是否循环
    /// </summary>
    bool Loop { get; set; }


    /// <summary>
    /// 重置切换效果
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置切换效果
    /// </summary>
    /// <param name="effectType">效果类型</param>
    /// <param name="duration">持续时间</param>
    /// <param name="speed">速度</param>
    void SetTransition(int effectType, float duration = 1.0f, int speed = 2);

    /// <summary>
    /// 设置切换声音
    /// </summary>
    /// <param name="soundFile">声音文件路径</param>
    /// <param name="loop">是否循环</param>
    void SetSound(string soundFile, bool loop = false);

    /// <summary>
    /// 设置定时
    /// </summary>
    /// <param name="advanceTime">前进时间</param>
    /// <param name="advanceOnTime">是否定时前进</param>
    void SetTiming(int advanceTime, bool advanceOnTime = true);

    /// <summary>
    /// 应用到指定幻灯片范围
    /// </summary>
    /// <param name="fromSlide">起始幻灯片</param>
    /// <param name="toSlide">结束幻灯片</param>
    void ApplyToRange(int fromSlide, int toSlide);

    /// <summary>
    /// 获取切换效果信息
    /// </summary>
    /// <returns>切换效果信息字符串</returns>
    string GetTransitionInfo();
}