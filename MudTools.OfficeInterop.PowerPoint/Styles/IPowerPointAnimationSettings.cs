//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 动画设置接口
/// </summary>
public interface IPowerPointAnimationSettings : IDisposable
{
    /// <summary>
    /// 获取或设置进入效果
    /// </summary>
    int EntryEffect { get; set; }

    /// <summary>
    /// 获取或设置动画顺序
    /// </summary>
    int AnimationOrder { get; set; }

    /// <summary>
    /// 获取或设置前进模式
    /// </summary>
    int AdvanceMode { get; set; }

    /// <summary>
    /// 获取或设置前进时间
    /// </summary>
    float AdvanceTime { get; set; }

    /// <summary>
    /// 获取声音效果
    /// </summary>
    IPowerPointSoundEffect SoundEffect { get; }

    /// <summary>
    /// 获取播放设置
    /// </summary>
    IPowerPointPlaySettings PlaySettings { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置是否动画背景
    /// </summary>
    bool AnimateBackground { get; set; }



    /// <summary>
    /// 播放动画
    /// </summary>
    /// <param name="from">起始时间</param>
    /// <param name="to">结束时间</param>
    /// <param name="repeats">重复次数</param>
    void Play(double from = 0, double to = 0, int repeats = 1);

    /// <summary>
    /// 停止动画
    /// </summary>
    void Stop();

    /// <summary>
    /// 暂停动画
    /// </summary>
    void Pause();

    /// <summary>
    /// 恢复动画
    /// </summary>
    void Resume();

    /// <summary>
    /// 重置动画设置
    /// </summary>
    void Reset();

    /// <summary>
    /// 应用动画方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    void ApplyAnimationScheme(int schemeIndex = -1);

    /// <summary>
    /// 设置动画参数
    /// </summary>
    /// <param name="entryEffect">进入效果</param>
    /// <param name="advanceMode">前进模式</param>
    /// <param name="advanceTime">前进时间</param>
    void SetAnimation(int entryEffect = 0, int advanceMode = 1, float advanceTime = 0);

    /// <summary>
    /// 获取动画设置信息
    /// </summary>
    /// <returns>动画设置信息字符串</returns>
    string GetAnimationSettingsInfo();
}
