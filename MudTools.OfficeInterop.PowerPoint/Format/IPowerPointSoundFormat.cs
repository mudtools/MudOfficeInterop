//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中的声音格式。
/// 提供播放、导入和导出声音文件的功能。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSoundFormat : IOfficeObject<IPowerPointSoundFormat, MsPowerPoint.SoundFormat>, IDisposable
{
    /// <summary>
    /// 播放当前声音。
    /// </summary>
    void Play();

    /// <summary>
    /// 从指定文件导入声音。
    /// </summary>
    /// <param name="fileName">要导入的声音文件的完整路径。</param>
    void Import(string fileName);

    /// <summary>
    /// 将当前声音导出到指定文件。
    /// </summary>
    /// <param name="fileName">导出文件的完整路径。</param>
    /// <returns>导出声音的格式类型。</returns>
    PpSoundFormatType? Export(string fileName);

    /// <summary>
    /// 获取当前声音的格式类型。
    /// </summary>
    PpSoundFormatType Type { get; }

    /// <summary>
    /// 获取声音源文件的完整路径。
    /// </summary>
    string SourceFullName { get; }
}