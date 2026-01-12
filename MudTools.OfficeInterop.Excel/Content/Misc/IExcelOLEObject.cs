//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表中的 ActiveX 控件或链接/嵌入的 OLE 对象。
/// OLEObject 对象是 OLEObjects 集合的成员。OLEObjects 集合包含单个工作表上的所有 OLE 对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelOLEObject : IOfficeObject<IExcelOLEObject, MsExcel.OLEObject>, IDisposable
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
    /// 获取表示对象右下角下方单元格的 Range 对象。只读。
    /// </summary>
    IExcelRange? BottomRightCell { get; }

    /// <summary>
    /// 将对象置于 Z 顺序的前面。
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将对象复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将所选对象以图片形式复制到剪贴板。
    /// </summary>
    /// <param name="appearance">可选项。XlPictureAppearance。指定应如何复制图片。</param>
    /// <param name="format">可选项。XlCopyPictureFormat。图片的格式。</param>
    void CopyPicture(XlPictureAppearance appearance = XlPictureAppearance.xlPrinter,
                     XlCopyPictureFormat format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 将对象剪切到剪贴板或将其粘贴到指定目标。
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除该对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制对象并返回对新副本的引用。
    /// </summary>
    /// <returns>复制的新对象。</returns>
    object? Duplicate();

    /// <summary>
    /// 获取或设置一个值，该值指示对象是否已启用。可读写布尔值。
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取或设置对象的高度（以磅为单位）。可读写 Double。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取对象在相似对象集合中的索引号。只读整数。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置对象左边缘到 A 列左边缘（在工作表上）或图表区左边缘（在图表上）的距离（以磅为单位）。可读写 Double。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示对象是否已锁定。如果工作表受保护，则锁定的对象无法修改。可读写布尔值。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置对象的名称。可读写字符串。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置单击指定对象时要运行的宏的名称。可读写字符串。
    /// </summary>
    string OnAction { get; set; }

    /// <summary>
    /// 获取或设置对象附加到其下方单元格的方式。可读写对象。
    /// </summary>
    object Placement { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示打印文档时是否打印该对象。可读写布尔值。
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 选择对象。
    /// </summary>
    /// <param name="replace">可选项。布尔值。True 表示用指定对象替换当前选定内容。False 表示扩展当前选定内容以包括任何先前选定的对象和指定对象。</param>
    void Select(bool? replace = null);

    /// <summary>
    /// 将对象发送到 Z 顺序的后面。
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 获取或设置对象顶部边缘到第 1 行顶部（在工作表上）或图表区顶部（在图表上）的距离（以磅为单位）。可读写 Double。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取表示指定对象左上角下方单元格的 Range 对象。只读。
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取或设置一个值，该值确定对象是否可见。可读写布尔值。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置对象的宽度（以磅为单位）。可读写 Double。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取对象的 Z 顺序位置。只读整数。
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取表示指定对象或对象的 ShapeRange 对象。只读。
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取表示对象边框的 Border 对象。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取表示指定对象内部的 Interior 对象。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示字体是否为阴影字体或对象是否有阴影。可读写布尔值。
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 激活该对象。
    /// </summary>
    void Activate();

    /// <summary>
    /// 获取或设置一个值，该值指示打开包含 OLE 对象的工作簿时是否自动加载该 OLE 对象。可读写布尔值。
    /// </summary>
    bool AutoLoad { get; set; }

    /// <summary>
    /// 获取一个值，该值指示源更改时是否自动更新 OLE 对象。
    /// 仅当对象为链接对象时有效（其 OLEType 属性必须为 xlOLELink）。只读布尔值。
    /// </summary>
    bool AutoUpdate { get; }

    /// <summary>
    /// 获取与此 OLE 对象关联的 OLE 自动化对象。只读对象。
    /// </summary>
    object Object { get; }

    /// <summary>
    /// 获取 OLE 对象类型。可以是以下 XlOLEType 常量之一：xlOLELink 或 xlOLEEmbed。
    /// 如果对象是链接的（存在于文件外部），则返回 xlOLELink；如果对象是嵌入的（完全包含在文件中），则返回 xlOLEEmbed。只读对象。
    /// </summary>
    XlOLEType OLEType { get; }

    /// <summary>
    /// 获取或设置指定对象的链接源名称。可读写字符串。
    /// </summary>
    string SourceName { get; set; }

    /// <summary>
    /// 更新链接。
    /// </summary>
    void Update();

    /// <summary>
    /// 向指定 OLE 对象的服务器发送动词。
    /// </summary>
    /// <param name="verb">可选项。XlOLEVerb。OLE 对象服务器应执行的动词。如果省略此参数，则发送默认动词。可用动词由对象的源应用程序决定。OLE 对象的典型动词是 Open 和 Primary（分别由 XlOLEVerb 常量 xlOpen 和 xlPrimary 表示）。</param>
    void Verb(XlOLEVerb verb = XlOLEVerb.xlVerbPrimary);

    /// <summary>
    /// 获取或设置链接到控件值的工作表区域。
    /// 如果在单元格中放置值，则控件会获取该值。同样，如果更改控件的值，则该值也会被放置到单元格中。可读写字符串。
    /// </summary>
    string LinkedCell { get; set; }

    /// <summary>
    /// 获取或设置用于填充指定列表框的工作表区域。
    /// 设置此属性会破坏列表框中任何现有的列表。可读写字符串。
    /// </summary>
    string ListFillRange { get; set; }

    /// <summary>
    /// 获取对象的编程标识符。只读字符串。
    /// </summary>
    string progID { get; }

    /// <summary>
    /// 获取或设置对象的备用 HTML。此属性保留供内部使用。
    /// </summary>
    string AltHTML { get; set; }
}