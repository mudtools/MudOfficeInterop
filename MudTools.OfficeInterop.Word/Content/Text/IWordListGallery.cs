//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示单个列表格式的库。
/// <para>注：ListGallery 对象是 ListGalleries 集合的成员。每个 ListGallery 对象表示“项目符号和编号”对话框中的三个选项卡之一。</para>
/// </summary>
public interface IWordListGallery : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取表示指定列表库的所有列表格式的 ListTemplates 集合。
    /// </summary>
    IWordListTemplates ListTemplates { get; }

    /// <summary>
    /// 获取一个值，该值指示指定的列表模板是否包含 Microsoft Word 中内置的格式。
    /// <para>注：索引号代表“项目符号和编号”对话框中指定选项卡上的列表格式。</para>
    /// </summary>
    /// <param name="index">列表格式的索引号（从 1 开始）。</param>
    /// <returns>如果列表模板已被修改则返回 true，否则返回 false。</returns>
    bool Modified(int index);

    /// <summary>
    /// 将列表库的指定列表模板重置为内置列表模板格式。
    /// </summary>
    /// <param name="index">要重置的列表模板的索引号（从 1 开始）。</param>
    void Reset(int index);
}