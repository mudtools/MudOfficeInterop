//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// PowerPoint 声音效果接口
/// </summary>
public interface IPowerPointSoundEffect : IDisposable
{
    /// <summary>
    /// 获取或设置声音名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 从文件导入声音
    /// </summary>
    /// <param name="fileName">文件路径</param>
    void ImportFromFile(string fileName);

    /// <summary>
    /// 播放声音
    /// </summary>
    void Play();

    /// <summary>
    /// 停止播放
    /// </summary>
    void Stop();

    /// <summary>
    /// 暂停播放
    /// </summary>
    void Pause();

    /// <summary>
    /// 恢复播放
    /// </summary>
    void Resume();

    /// <summary>
    /// 删除声音效果
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取声音效果信息
    /// </summary>
    /// <returns>声音效果信息字符串</returns>
    string GetSoundEffectInfo();
}