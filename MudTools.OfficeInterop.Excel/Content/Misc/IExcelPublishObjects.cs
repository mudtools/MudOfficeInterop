//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作簿中所有PublishObject对象的集合。每个PublishObject对象表示已保存到网页的工作簿项，可以根据对象的属性和方法指定的值进行刷新。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelPublishObjects : IEnumerable<IExcelPublishObject?>, IOfficeObject<IExcelPublishObjects>, IDisposable
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
    /// 获取对象数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取对象
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    IExcelPublishObject? this[int index] { get; }

    /// <summary>
    /// 获取对象
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    IExcelPublishObject? this[string name] { get; }

    /// <summary>
    /// 创建一个表示已保存到网页的文档项的对象。这样的对象有助于在Excel中对文档进行自动化更改时，随后更新网页。
    /// </summary>
    /// <param name="sourceType">必需。源类型。</param>
    /// <param name="filename">必需。源对象保存到的URL（在intranet或Web上）或路径（本地或网络）。</param>
    /// <param name="sheet">可选。保存为网页的工作表名称。</param>
    /// <param name="source">可选。用于标识源类型为以下常量之一的项的唯一名称：xlSourceAutoFilter、xlSourceChart、xlSourcePivotTable、xlSourcePrintArea、xlSourceQuery或xlSourceRange。如果SourceType是xlSourceRange，则Source指定一个范围（可以是定义的名称）。如果SourceType是xlSourceChart、xlSourcePivotTable或xlSourceQuery，则Source指定图表、数据透视表或查询表的名称。</param>
    /// <param name="htmlType">可选。指定项是否保存为交互式Microsoft Office Web组件或静态文本和图像。</param>
    /// <param name="divID">可选。HTML DIV标记中用于标识网页上项的唯一标识符。</param>
    /// <param name="title">可选。网页的标题。</param>
    /// <returns>新创建的PublishObject对象。</returns>
    IExcelPublishObject? Add(XlSourceType sourceType, string filename,
                            string? sheet = null, string? source = null,
                            XlHtmlType? htmlType = null, string? divID = null, string? title = null);

    /// <summary>
    /// 删除整个集合中的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将文档中的项或项集合保存到网页。
    /// </summary>
    void Publish();
}