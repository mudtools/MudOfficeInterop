//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// Word 书签接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordBookmark : IOfficeObject<IWordBookmark>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取指定书签对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取表示指定书签对象中包含的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取一个值，该值指示指定书签是否为空。
    /// </summary>
    bool Empty { get; }

    /// <summary>
    /// 获取或设置书签的起始字符位置。
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置书签的结束字符位置。
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取一个值，该值指示指定书签是否为表格列。
    /// </summary>
    bool Column { get; }

    /// <summary>
    /// 获取指定书签的故事类型。
    /// </summary>
    WdStoryType StoryType { get; }

    /// <summary>
    /// 选择指定的书签对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除指定的书签对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将标识为 Name 参数的新书签设置为由指定书签对象标记的位置。
    /// </summary>
    /// <param name="name">新书签的名称。</param>
    /// <returns>复制的新书签对象。</returns>
    IWordBookmark? Copy(string name);
}
