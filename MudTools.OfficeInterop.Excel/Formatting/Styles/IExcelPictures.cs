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
public interface IExcelPictures : IEnumerable<IExcelPicture>, IDisposable
{
    #region 基础属性
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
    IExcelPicture? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定图片对象。
    /// 若名称为空、底层对象为空或访问异常，返回 null。
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>对应的图片对象，或 null</returns>
    IExcelPicture? this[string name] { get; }

    /// <summary>
    /// 获取图片集合所属的父对象（如 Worksheet）。
    /// 若底层对象无效，返回 null。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置图片集合是否启用（可交互）。
    /// 若底层对象无效或值为 null，不执行设置操作。
    /// </summary>
    bool? Enabled { get; set; }

    /// <summary>
    /// 获取或设置图片集合是否可见。
    /// 若底层对象无效或值为 null，不执行设置操作。
    /// </summary>
    bool? Visible { get; set; }

    /// <summary>
    /// 获取图片集合在 Z 轴上的层叠顺序（只读）。
    /// 若底层对象无效，返回 null。
    /// </summary>
    int? ZOrder { get; }

    /// <summary>
    /// 获取或设置图片集合是否锁定（防止用户修改）。
    /// 若底层对象无效或值为 null，不执行设置操作。
    /// </summary>
    bool? Locked { get; set; }

    /// <summary>
    /// 获取或设置图片集合是否显示阴影效果。
    /// 若底层对象无效或值为 null，不执行设置操作。
    /// </summary>
    bool? Shadow { get; set; }

    /// <summary>
    /// 获取或设置图片集合是否随工作表打印。
    /// 若底层对象无效或值为 null，不执行设置操作。
    /// </summary>
    bool? PrintObject { get; set; }

    /// <summary>
    /// 获取或设置点击图片时执行的宏名称。
    /// 若底层对象无效，不执行设置操作。
    /// </summary>
    string? OnAction { get; set; }

    /// <summary>
    /// 获取或设置图片关联的公式（如链接到单元格）。
    /// 若底层对象无效，不执行设置操作。
    /// </summary>
    string? Formula { get; set; }

    /// <summary>
    /// 获取或设置图片相对于单元格的定位方式（自由浮动、随单元格移动等）。
    /// 若底层对象无效，不执行设置操作。
    /// </summary>
    XlPlacement Placeholder { get; set; }

    /// <summary>
    /// 获取或设置图片高度（单位：磅）。
    /// 若底层对象无效，不执行设置操作；获取时若无效返回 0。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置图片宽度（单位：磅）。
    /// 若底层对象无效，不执行设置操作；获取时若无效返回 0。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置图片左边缘位置（单位：磅）。
    /// 若底层对象无效，不执行设置操作；获取时若无效返回 0。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图片上边缘位置（单位：磅）。
    /// 若底层对象无效，不执行设置操作；获取时若无效返回 0。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取图片集合对应的 ShapeRange 对象（用于批量操作形状）。
    /// 若底层对象无效，返回 null。
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取图片集合的边框样式对象。
    /// 若底层对象无效，返回 null。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取图片集合的内部填充样式对象。
    /// 若底层对象无效，返回 null。
    /// </summary>
    IExcelInterior? Interior { get; }
    #endregion

    #region 创建和添加

    /// <summary>
    /// 从指定文件路径插入一张新图片。
    /// 文件不存在或插入失败时返回 null，不抛出异常。
    /// </summary>
    /// <param name="filename">图片文件的完整路径</param>
    /// <param name="Converter">保留参数，通常传 null</param>
    /// <returns>新创建的图片对象，或 null</returns>
    IExcelPicture? Insert(string filename, object Converter);

    /// <summary>
    /// 在指定位置添加一个空图片占位符。
    /// </summary>
    /// <param name="Left">左边缘位置（单位：磅）</param>
    /// <param name="Top">上边缘位置（单位：磅）</param>
    /// <param name="Width">宽度（单位：磅）</param>
    /// <param name="Height">高度（单位：磅）</param>
    /// <returns>新创建的图片对象，或 null</returns>
    IExcelPicture? Add(double Left, double Top, double Width, double Height);

    /// <summary>
    /// 将当前图片集合中的所有图片组合成一个组对象。
    /// </summary>
    /// <returns>组合后的组对象，或 null</returns>
    IExcelGroupObject? Group();

    /// <summary>
    /// 从字节数组插入图片（支持内存中图片数据）。
    /// 自动创建并清理临时文件，失败时返回 null，不抛出异常。
    /// </summary>
    /// <param name="imageBytes">图片的字节数组</param>
    /// <param name="imageFormat">图片格式扩展名，如 "png"、"jpg"（默认 "png"）</param>
    /// <param name="left">左边缘位置（单位：磅，默认 0）</param>
    /// <param name="top">上边缘位置（单位：磅，默认 0）</param>
    /// <param name="width">宽度（单位：磅，默认 -1 表示原始尺寸）</param>
    /// <param name="height">高度（单位：磅，默认 -1 表示原始尺寸）</param>
    /// <returns>新创建的图片对象，或 null</returns>
    IExcelPicture? AddFromBytes(byte[] imageBytes, string imageFormat = "png",
                              double left = 0, double top = 0, double width = -1, double height = -1);

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称模糊匹配查找图片（支持包含关系）。
    /// </summary>
    /// <param name="name">要匹配的名称片段</param>
    /// <returns>匹配的图片数组，无匹配时返回空数组</returns>
    IExcelPicture[] FindByName(string name);

    /// <summary>
    /// 根据位置查找图片（支持容差）。
    /// </summary>
    /// <param name="left">目标左边缘位置</param>
    /// <param name="top">目标上边缘位置</param>
    /// <param name="tolerance">允许的位置误差（默认 10 磅）</param>
    /// <returns>匹配的图片数组，无匹配时返回空数组</returns>
    IExcelPicture[] FindByPosition(double left, double top, double tolerance = 10);

    /// <summary>
    /// 根据尺寸查找图片（支持容差）。
    /// </summary>
    /// <param name="width">目标宽度</param>
    /// <param name="height">目标高度</param>
    /// <param name="tolerance">允许的尺寸误差（默认 10 磅）</param>
    /// <returns>匹配的图片数组，无匹配时返回空数组</returns>
    IExcelPicture[] FindBySize(double width, double height, double tolerance = 10);

    /// <summary>
    /// 获取所有当前可见的图片。
    /// </summary>
    /// <returns>可见图片数组，无可见图片时返回空数组</returns>
    IExcelPicture[] GetVisiblePictures();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除集合中所有图片（从后往前删除，避免索引错乱）。
    /// 删除过程中发生异常会被记录但不中断流程。
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引位置的图片。
    /// 索引无效或删除失败时静默忽略，不抛出异常。
    /// </summary>
    /// <param name="index">要删除的图片索引（从 1 开始）</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的图片对象（当前实现可能存在问题，建议使用索引或名称删除）。
    /// </summary>
    /// <param name="picture">要删除的图片对象</param>
    void Delete(IExcelPicture picture);

    /// <summary>
    /// 批量删除多个指定索引的图片（按降序删除避免索引漂移）。
    /// </summary>
    /// <param name="indices">要删除的图片索引数组</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 将所有图片置于最上层（Z-Order 最前）。
    /// 操作失败时记录警告，不抛出异常。
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将所有图片置于最下层（Z-Order 最后）。
    /// 操作失败时记录警告，不抛出异常。
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 复制当前选中的图片到剪贴板。
    /// 操作失败时记录警告，不抛出异常。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切当前选中的图片到剪贴板。
    /// 操作失败时记录警告，不抛出异常。
    /// </summary>
    void Cut();

    /// <summary>
    /// 以指定外观和格式复制图片到剪贴板（常用于粘贴为图片而非对象）。
    /// </summary>
    /// <param name="Appearance">复制时的显示外观（默认打印机效果）</param>
    /// <param name="Format">复制格式（默认位图图片）</param>
    void CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter,
                     XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 复制当前图片并粘贴到同一位置（创建副本）。
    /// 操作失败时记录警告，不抛出异常。
    /// </summary>
    void Duplicate();

    #endregion
}
