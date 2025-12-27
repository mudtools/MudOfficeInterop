//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示已保存到网页的工作簿项，可以根据PublishObject对象的属性和方法指定的值进行刷新。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPublishObject : IOfficeObject<IExcelPublishObject>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 删除此发布对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将文档中的项或项集合保存到网页。
    /// </summary>
    /// <param name="create">可选。如果HTML文件存在，将此参数设置为True将替换文件，设置为False将在文件末尾插入项。如果文件不存在，无论Create参数的值如何，都会创建文件。</param>
    void Publish(bool? create = null);

    /// <summary>
    /// 获取用于标识网页上HTML &lt;DIV&gt;标记的唯一标识符。
    /// </summary>
    string DivID { get; }

    /// <summary>
    /// 获取指定PublishObject对象的工作表名称。
    /// </summary>
    string Sheet { get; }

    /// <summary>
    /// 获取标识正在发布的项类型的值。
    /// </summary>
    XlSourceType SourceType { get; }

    /// <summary>
    /// 获取唯一名称，用于标识SourceType属性值为xlSourceRange、xlSourceChart、xlSourcePrintArea、xlSourceAutoFilter、xlSourcePivotTable或xlSourceQuery的项。
    /// </summary>
    string Source { get; }

    /// <summary>
    /// 获取或设置将指定项保存到网页时Excel生成的HTML类型。
    /// </summary>
    XlHtmlType HtmlType { get; set; }

    /// <summary>
    /// 获取或设置文档保存为网页时网页的标题。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置指定源对象保存位置（在intranet或Web上）的URL或路径（本地或网络）。
    /// </summary>
    string Filename { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示保存工作簿时，如果PublishObjects集合中的任何项的AutoRepublish属性设置为True，则重新发布它。默认值为False。
    /// </summary>
    bool AutoRepublish { get; set; }
}