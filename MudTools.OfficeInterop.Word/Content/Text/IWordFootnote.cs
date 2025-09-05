//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 脚注的封装接口。
/// </summary>
public interface IWordFootnote : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取脚注索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置引用文本范围。
    /// </summary>
    IWordRange Reference { get; }

    /// <summary>
    /// 获取或设置脚注文本范围。
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取脚注编号。
    /// </summary>
    string Number { get; }

    /// <summary>
    /// 获取或设置脚注字体。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取或设置脚注段落格式。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// 选择脚注。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除脚注。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制脚注。
    /// </summary>
    /// <returns>复制的脚注。</returns>
    IWordFootnote Copy();

    /// <summary>
    /// 更新脚注编号。
    /// </summary>
    void Update();

    /// <summary>
    /// 获取脚注引用位置。
    /// </summary>
    /// <returns>引用位置范围。</returns>
    IWordRange GetReferenceRange();

    /// <summary>
    /// 获取脚注内容位置。
    /// </summary>
    /// <returns>内容位置范围。</returns>
    IWordRange GetContentRange();

    /// <summary>
    /// 修改脚注文本内容。
    /// </summary>
    /// <param name="newText">新文本内容。</param>
    void ModifyText(string newText);

    /// <summary>
    /// 获取脚注文本内容。
    /// </summary>
    /// <returns>脚注文本。</returns>
    string GetText();

    /// <summary>
    /// 设置脚注文本内容。
    /// </summary>
    /// <param name="text">文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 检查脚注是否包含指定文本。
    /// </summary>
    /// <param name="text">要检查的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <returns>是否包含。</returns>
    bool ContainsText(string text, bool matchCase = false);
}