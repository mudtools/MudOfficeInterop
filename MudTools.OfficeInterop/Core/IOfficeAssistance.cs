//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中帮助系统接口的封装。
/// 提供对 Office 帮助系统的访问和控制功能。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore", ComClassName = "IAssistance")]
public interface IOfficeAssistance : IDisposable
{
    /// <summary>
    /// 搜索帮助内容。
    /// </summary>
    /// <param name="query">搜索查询字符串。</param>
    /// <param name="scope">帮助搜索范围。</param>
    void SearchHelp(string query, string scope = "");

    /// <summary>
    /// 显示上下文相关的帮助。
    /// </summary>
    /// <param name="helpId">帮助标识符。</param>
    /// <param name="scope">帮助显示范围。</param>
    void ShowHelp(string helpId = "", string scope = "");


    /// <summary>
    /// 设置默认的帮助上下文。
    /// </summary>
    /// <param name="helpId">帮助标识符。</param>
    void SetDefaultContext(string helpId);

    /// <summary>
    /// 清除指定的帮助上下文。
    /// </summary>
    /// <param name="helpId">要清除的帮助标识符。</param>
    void ClearDefaultContext(string helpId);
}