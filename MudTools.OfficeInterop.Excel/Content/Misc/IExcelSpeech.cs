//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 包含与语音相关的方法和属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSpeech : IOfficeObject<IExcelSpeech, MsExcel.Speech>, IDisposable
{
    /// <summary>
    /// 播放指定的文本字符串。
    /// </summary>
    /// <param name="text">要朗读的文本。</param>
    /// <param name="speakAsync">如果为 true，则文本将异步朗读（方法不会等待文本朗读完毕）。如果为 false，则文本将同步朗读（方法等待文本朗读完毕后再继续）。默认为 false。</param>
    /// <param name="speakXml">如果为 true，则文本将被解释为 XML。如果为 false，则文本不会被解释为 XML，因此任何 XML 标签将被读取而不被解释。默认为 false。</param>
    /// <param name="purge">如果为 true，则在朗读文本之前终止当前语音并清除任何缓冲的文本。如果为 false，则不会终止当前语音，也不会在朗读文本之前清除缓冲的文本。默认为 false。</param>
    void Speak(string text, bool? speakAsync = null, bool? speakXml = null, bool? purge = null);

    /// <summary>
    /// 获取或设置朗读单元格的顺序。
    /// </summary>
    XlSpeakDirection Direction { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在按下 ENTER 键或活动单元格编辑完成时是否朗读活动单元格。将此属性设置为 true 将启用此模式。false 将关闭此模式。
    /// </summary>
    bool SpeakCellOnEnter { get; set; }
}