//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个超链接（Hyperlink）的封装接口。
/// </summary>
public interface IWordHyperlink : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置超链接的显示文本。
    /// </summary>
    string TextToDisplay { get; set; }

    /// <summary>
    /// 获取或设置超链接的目标地址。
    /// </summary>
    string Address { get; set; }

    /// <summary>
    /// 获取或设置超链接的子地址（如书签名称）。
    /// </summary>
    string SubAddress { get; set; }

    /// <summary>
    /// 获取或设置超链接的屏幕提示文本。
    /// </summary>
    string ScreenTip { get; set; }

    /// <summary>
    /// 获取超链接所在的范围。
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取超链接的类型。
    /// </summary>
    MsoHyperlinkType Type { get; }

    /// <summary>
    /// 获取或设置是否在新窗口中打开超链接。
    /// </summary>
    string Target { get; set; }

    /// <summary>
    /// 获取或设置超链接的邮件主题（仅适用于邮件链接）。
    /// </summary>
    string EmailSubject { get; set; }

    /// <summary>
    /// 删除此超链接。
    /// </summary>
    void Delete();

    /// <summary>
    /// 跟随此超链接。
    /// </summary>
    void Follow();
}