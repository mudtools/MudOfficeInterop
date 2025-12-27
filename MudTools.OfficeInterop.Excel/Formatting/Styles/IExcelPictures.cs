//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Pictures 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Pictures 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelPictures : IOfficeObject<IExcelPictures>, IEnumerable<IExcelPicture?>, IDisposable
{
    /// <summary>
    /// 获取当前图片集合中包含的图片总数。
    /// 若底层对象已被释放或无效，返回 0。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定位置的图片对象（索引从 1 开始）。
    /// 若索引无效、底层对象为空或访问异常，返回 null。
    /// </summary>
    /// <param name="index">图片索引，从 1 开始</param>
    /// <returns>对应的图片对象，或 null</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPicture? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定图片对象。
    /// 若名称为空、底层对象为空或访问异常，返回 null。
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>对应的图片对象，或 null</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelPicture? this[string name] { get; }

    /// <summary>
    /// 获取图片集合所属的父对象（如 Worksheet）。
    /// 若底层对象无效，返回 null。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 将图片集合中的所有图片置于所有其他对象的前面。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? BringToFront();

    /// <summary>
    /// 复制图片集合中的所有图片。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? Copy();

    /// <summary>
    /// 将图片集合中的所有图片作为图片复制到剪贴板。
    /// </summary>
    /// <param name="Appearance">指定图片的外观。</param>
    /// <param name="Format">指定图片的格式。</param>
    /// <returns>操作结果对象。</returns>
    object? CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter, XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 剪切图片集合中的所有图片。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? Cut();

    /// <summary>
    /// 删除图片集合中的所有图片。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? Delete();

    /// <summary>
    /// 复制图片集合中的所有图片并返回新对象。
    /// </summary>
    /// <returns>复制的对象。</returns>
    object? Duplicate();

    /// <summary>
    /// 获取或设置图片集合是否可用。
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取或设置图片集合的高度（以磅为单位）。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置图片集合左边缘到其父对象左边缘的距离（以磅为单位）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图片集合是否被锁定。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置当用户单击图片集合时要运行的过程的名称。
    /// </summary>
    string OnAction { get; set; }

    /// <summary>
    /// 获取或设置图片集合附加到其下方单元格的方式。
    /// </summary>
    object Placement { get; set; }

    /// <summary>
    /// 获取或设置打印包含图片集合的工作表时是否打印该集合。
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 选择图片集合。
    /// </summary>
    /// <param name="replace">如果为True，则用当前选择替换当前选择；如果为False，则将当前选择扩展到包括以前选择的对象。</param>
    /// <returns>操作结果对象。</returns>
    object? Select(bool? replace = null);

    /// <summary>
    /// 将图片集合中的所有图片置于所有其他对象的后面。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object? SendToBack();

    /// <summary>
    /// 获取或设置图片集合上边缘到其父对象上边缘的距离（以磅为单位）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置图片集合是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置图片集合的宽度（以磅为单位）。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取图片集合在z-order中的位置。
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取表示图片集合中所有图片的形状范围。
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取图片集合的边框。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取图片集合的内部区域。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置图片集合是否有阴影。
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置图片集合的公式。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 向集合中添加一张图片。
    /// </summary>
    /// <param name="Left">新图片的左边缘位置（以磅为单位）。</param>
    /// <param name="Top">新图片的上边缘位置（以磅为单位）。</param>
    /// <param name="Width">新图片的宽度（以磅为单位）。</param>
    /// <param name="Height">新图片的高度（以磅为单位）。</param>
    /// <returns>新添加的图片对象。</returns>
    IExcelPicture? Add(double Left, double Top, double Width, double Height);

    /// <summary>
    /// 将图片集合中的所有图片组合成一个组。
    /// </summary>
    /// <returns>表示组合对象的新对象。</returns>
    IExcelGroupObject? Group();

    /// <summary>
    /// 从文件中插入一张图片。
    /// </summary>
    /// <param name="Filename">要插入的图片的文件名。</param>
    /// <param name="Converter">用于转换图片的转换器。</param>
    /// <returns>新插入的图片对象。</returns>
    IExcelPicture? Insert(string Filename, object? Converter = null);

    /// <summary>
    /// 从剪贴板粘贴一张图片。
    /// </summary>
    /// <param name="Link">如果为True，则图片将链接到其源文件。</param>
    /// <returns>新粘贴的图片对象。</returns>
    IExcelPicture? Paste(bool? Link = null);
}
