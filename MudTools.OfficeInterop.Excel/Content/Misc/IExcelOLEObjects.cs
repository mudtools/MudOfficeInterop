//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示指定工作表上所有 OLEObject 对象的集合。每个 OLEObject 对象表示一个 ActiveX 控件或链接/嵌入的 OLE 对象。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelOLEObjects : IEnumerable<IExcelOLEObject?>, IOfficeObject<IExcelOLEObjects>, IDisposable
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
    [ComPropertyWrap(NeedConvert = true)]
    IExcelOLEObject? this[int index] { get; }

    /// <summary>
    /// 获取对象
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelOLEObject? this[string name] { get; }


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
    void CopyPicture(XlPictureAppearance appearance = XlPictureAppearance.xlPrinter, XlCopyPictureFormat format = XlCopyPictureFormat.xlPicture);

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
    /// 获取或设置一个值，该值指示对象是否已启用。
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取或设置对象的高度（以磅为单位）。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置对象左边缘到 A 列左边缘（在工作表上）或图表区左边缘（在图表上）的距离（以磅为单位）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示对象是否已锁定。如果工作表受保护，则锁定的对象无法修改。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置单击指定对象时要运行的宏的名称。
    /// </summary>
    string OnAction { get; set; }

    /// <summary>
    /// 获取或设置对象附加到其下方单元格的方式。
    /// </summary>
    object Placement { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示打印文档时是否打印该对象。
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
    /// 获取或设置对象顶部边缘到第 1 行顶部（在工作表上）或图表区顶部（在图表上）的距离（以磅为单位）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值确定对象是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置对象的宽度（以磅为单位）。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取对象的 Z 顺序位置。
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取表示指定对象或对象的 ShapeRange 对象。
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
    /// 获取或设置一个值，该值指示字体是否为阴影字体或对象是否有阴影。
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示打开包含 OLE 对象的工作簿时是否自动加载该 OLE 对象。
    /// </summary>
    bool AutoLoad { get; set; }

    /// <summary>
    /// 获取或设置指定对象的链接源名称。
    /// </summary>
    string SourceName { get; set; }

    /// <summary>
    /// 向工作表添加新的 OLE 对象。返回 OLEObject 对象。
    /// </summary>
    /// <param name="classType">可选项。（必须指定 ClassType 或 FileName。）包含要创建的对象的编程标识符的字符串。如果指定了 ClassType，则忽略 FileName 和 Link 参数。</param>
    /// <param name="filename">可选项。（必须指定 ClassType 或 FileName。）指定用于创建 OLE 对象的文件的字符串。</param>
    /// <param name="link">可选项。True 表示使基于 FileName 的新 OLE 对象链接到该文件。如果对象未链接，则对象将创建为文件的副本。默认值为 False。</param>
    /// <param name="displayAsIcon">可选项。True 表示将新 OLE 对象显示为图标或其常规图片。如果此参数为 True，则可以使用 IconFileName 和 IconIndex 来指定图标。</param>
    /// <param name="iconFileName">可选项。字符串，指定包含要显示的图标的文件。仅当 DisplayAsIcon 为 True 时才使用此参数。如果未指定此参数或文件不包含图标，则使用 OLE 类的默认图标。</param>
    /// <param name="iconIndex">可选项。图标文件中的图标编号。仅当 DisplayAsIcon 为 True 且 IconFileName 引用包含图标的有效文件时才使用此参数。如果 IconFileName 指定的文件中不存在具有给定索引号的图标，则使用文件中的第一个图标。</param>
    /// <param name="iconLabel">可选项。字符串，指定要在图标下方显示的标签。仅当 DisplayAsIcon 为 True 时才使用此参数。如果省略此参数或为空字符串 ("")，则不显示标题。</param>
    /// <param name="left">可选项。新对象的初始坐标（以磅为单位），相对于工作表上单元格 A1 的左上角或图表的左上角。</param>
    /// <param name="top">可选项。新对象的初始坐标（以磅为单位），相对于工作表上单元格 A1 的左上角或图表的左上角。</param>
    /// <param name="width">可选项。新对象的初始大小（以磅为单位）。</param>
    /// <param name="height">可选项。新对象的初始大小（以磅为单位）。</param>
    /// <returns>新创建的 OLEObject 对象。</returns>
    IExcelOLEObject? Add(string? classType = null, string? filename = null,
                        bool? link = null, bool? displayAsIcon = null,
                        string? iconFileName = null, int? iconIndex = null,
                        string? iconLabel = null, double? left = null,
                        double? top = null, double? width = null,
                        double? height = null);
}