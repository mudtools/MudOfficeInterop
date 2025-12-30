//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 批注的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordComment : IOfficeObject<IWordComment>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示指定对象中包含的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取表示脚注、尾注或批注引用标记的 Range 对象。
    /// </summary>
    IWordRange? Reference { get; }

    /// <summary>
    /// 获取表示指定批注标记的文本范围的 Range 对象。
    /// </summary>
    IWordRange? Scope { get; }

    /// <summary>
    /// 获取表示集合中项目位置的整数。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置批注的作者名称。
    /// </summary>
    string Author { get; set; }

    /// <summary>
    /// 获取或设置与特定批注关联的用户缩写。
    /// </summary>
    string Initial { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示与批注关联的文本是否在屏幕提示中显示。
    /// </summary>
    bool ShowTip { get; set; }

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 打开指定的批注进行编辑。
    /// </summary>
    void Edit();

    /// <summary>
    /// 获取输入批注的日期和时间。
    /// </summary>
    DateTime Date { get; }

    /// <summary>
    /// 获取一个布尔值，该值表示批注是否为手写批注。
    /// </summary>
    bool IsInk { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示批注是否已完成。
    /// </summary>
    bool Done { get; set; }

    /// <summary>
    /// 获取此批注的祖先批注。
    /// </summary>
    IWordComment? Ancestor { get; }

    /// <summary>
    /// 获取与此批注关联的共同作者联系人。
    /// </summary>
    IWordCoAuthor? Contact { get; }

    /// <summary>
    /// 递归删除此批注及其所有回复。
    /// </summary>
    void DeleteRecursively();

    /// <summary>
    /// 获取此批注的所有回复的集合。
    /// </summary>
    IWordComments? Replies { get; }
}