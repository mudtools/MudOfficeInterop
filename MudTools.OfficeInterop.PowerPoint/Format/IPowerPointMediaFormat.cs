//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示媒体的格式设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointMediaFormat : IDisposable
{
    /// <summary>
    /// 获取创建此媒体格式对象的应用程序。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取媒体格式对象的父对象。
    /// </summary>
    /// <value>父对象，通常是媒体对象所属的形状或幻灯片。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置媒体的音量。
    /// </summary>
    /// <value>媒体音量，范围通常为 0.0（静音）到 1.0（最大）。</value>
    float Volume { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示媒体是否静音。
    /// </summary>
    /// <value>如果媒体静音，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool Muted { get; set; }

    /// <summary>
    /// 获取媒体的总长度（以毫秒为单位）。
    /// </summary>
    /// <value>媒体的总时长（毫秒）。</value>
    int Length { get; }

    /// <summary>
    /// 获取或设置媒体播放的起始点（以毫秒为单位）。
    /// </summary>
    /// <value>起始点（毫秒）。</value>
    int StartPoint { get; set; }

    /// <summary>
    /// 获取或设置媒体播放的结束点（以毫秒为单位）。
    /// </summary>
    /// <value>结束点（毫秒）。</value>
    int EndPoint { get; set; }

    /// <summary>
    /// 获取或设置媒体淡入的持续时间（以毫秒为单位）。
    /// </summary>
    /// <value>淡入持续时间（毫秒）。</value>
    int FadeInDuration { get; set; }

    /// <summary>
    /// 获取或设置媒体淡出的持续时间（以毫秒为单位）。
    /// </summary>
    /// <value>淡出持续时间（毫秒）。</value>
    int FadeOutDuration { get; set; }

    /// <summary>
    /// 获取媒体的书签集合。
    /// </summary>
    /// <value>媒体书签的集合。</value>
    [ComPropertyWrap(ComNamespace = "MsPowerPoint")]
    IPowerPointMediaBookmarks? MediaBookmarks { get; }

    /// <summary>
    /// 将指定时间点的媒体帧设置为显示图片。
    /// </summary>
    /// <param name="position">用作显示图片的媒体时间点（毫秒）。</param>
    void SetDisplayPicture(int position);

    /// <summary>
    /// 从指定文件设置媒体的显示图片。
    /// </summary>
    /// <param name="filePath">包含要使用的图片的文件路径。</param>
    void SetDisplayPictureFromFile(string filePath);

    /// <summary>
    /// 根据指定参数对媒体进行重新采样。
    /// </summary>
    /// <param name="trim">一个值，指示是否在重新采样前修剪媒体。</param>
    /// <param name="sampleHeight">目标采样高度（像素）。</param>
    /// <param name="sampleWidth">目标采样宽度（像素）。</param>
    /// <param name="videoFrameRate">目标视频帧率（帧/秒）。</param>
    /// <param name="audioSamplingRate">目标音频采样率（Hz）。</param>
    /// <param name="videoBitRate">目标视频比特率（比特/秒）。</param>
    void Resample(bool trim = false, int sampleHeight = 768, int sampleWidth = 1280, int videoFrameRate = 24, int audioSamplingRate = 48000, int videoBitRate = 7000000);

    /// <summary>
    /// 使用预定义的媒体配置文件对媒体进行重新采样。
    /// </summary>
    /// <param name="profile">指定要使用的预定义媒体配置文件。</param>
    void ResampleFromProfile(PpResampleMediaProfile profile = PpResampleMediaProfile.ppResampleMediaProfileSmall);

    /// <summary>
    /// 获取当前重新采样操作的状态。
    /// </summary>
    /// <value>重新采样任务的状态。</value>
    PpMediaTaskStatus ResamplingStatus { get; }

    /// <summary>
    /// 获取一个值，指示媒体是否为链接文件。
    /// </summary>
    /// <value>如果媒体链接到外部文件，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool IsLinked { get; }

    /// <summary>
    /// 获取一个值，指示媒体是否已嵌入到演示文稿中。
    /// </summary>
    /// <value>如果媒体已嵌入，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    bool IsEmbedded { get; }

    /// <summary>
    /// 获取媒体的音频采样率。
    /// </summary>
    /// <value>音频采样率（Hz）。</value>
    int AudioSamplingRate { get; }

    /// <summary>
    /// 获取媒体的视频帧率。
    /// </summary>
    /// <value>视频帧率（帧/秒）。</value>
    int VideoFrameRate { get; }

    /// <summary>
    /// 获取媒体采样的高度。
    /// </summary>
    /// <value>采样高度（像素）。</value>
    int SampleHeight { get; }

    /// <summary>
    /// 获取媒体采样的宽度。
    /// </summary>
    /// <value>采样宽度（像素）。</value>
    int SampleWidth { get; }

    /// <summary>
    /// 获取媒体使用的视频压缩类型。
    /// </summary>
    /// <value>视频压缩编解码器的名称。</value>
    string? VideoCompressionType { get; }

    /// <summary>
    /// 获取媒体使用的音频压缩类型。
    /// </summary>
    /// <value>音频压缩编解码器的名称。</value>
    string? AudioCompressionType { get; }
}