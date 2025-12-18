//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中一个组合形状（Grouped Shape）的封装接口。
/// 对应 COM 对象：实际上是 Microsoft.Office.Interop.Excel.Shape（Type = msoGroup）
/// 本接口语义命名为 GroupObject，用于管理组合形状及其子项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelGroupObject : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet 或 Shapes 集合）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置组合形状的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取组合形状的左上角单元格。
    /// 返回封装后的 <see cref="IExcelRange"/>。
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取组合形状的右下角单元格。
    /// 返回封装后的 <see cref="IExcelRange"/>。
    /// </summary>
    IExcelRange? BottomRightCell { get; }

    /// <summary>
    /// 获取组合形状的形状区域。
    /// 返回封装后的 <see cref="IExcelShapeRange"/>。
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取组合形状的边框格式。
    /// 返回封装后的 <see cref="IExcelBorder"/>。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取组合形状的字体格式。
    /// 返回封装后的 <see cref="IExcelFont"/>。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取组合形状的内部填充格式。
    /// 返回封装后的 <see cref="IExcelInterior"/>。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置是否自动缩进。
    /// 当设置为 true 时，文本会根据需要自动添加缩进。
    /// </summary>
    bool? AddIndent { get; set; }

    /// <summary>
    /// 获取或设置是否自动调整大小。
    /// 当设置为 true 时，对象会根据内容自动调整大小。
    /// </summary>
    bool? AutoSize { get; set; }

    /// <summary>
    /// 获取或设置对象是否可见。
    /// true 表示可见，false 表示隐藏。
    /// </summary>
    bool? Visible { get; set; }

    /// <summary>
    /// 获取或设置是否具有圆角。
    /// true 表示对象的角是圆角，false 表示是直角。
    /// </summary>
    bool? RoundedCorners { get; set; }

    /// <summary>
    /// 获取或设置是否具有阴影效果。
    /// true 表示显示阴影，false 表示不显示阴影。
    /// </summary>
    bool? Shadow { get; set; }

    /// <summary>
    /// 获取对象的 Z 轴顺序。
    /// 数值越大表示越靠近前景，数值越小表示越靠近背景。
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取或设置阅读顺序。
    /// 用于指定文本的阅读方向（从左到右或从右到左）。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置箭头头部长度。
    /// 指定形状线条末端箭头的长度。
    /// </summary>
    XlArrowHeadLength ArrowHeadLength { get; set; }

    /// <summary>
    /// 获取或设置箭头头部样式。
    /// 指定形状线条末端箭头的样式类型。
    /// </summary>
    XlArrowHeadStyle ArrowHeadStyle { get; set; }

    /// <summary>
    /// 获取或设置箭头头部宽度。
    /// 指定形状线条末端箭头的宽度。
    /// </summary>
    XlArrowHeadWidth ArrowHeadWidth { get; set; }

    /// <summary>
    /// 获取或设置对象的方向。
    /// 指定对象（如文本）的排列方向。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取组合形状的左边缘位置（单位：磅）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取组合形状的上边缘位置（单位：磅）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置组合形状的宽度（单位：磅）。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置组合形状的高度（单位：磅）。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 取消组合，将组合形状拆分为独立的子形状。
    /// 返回新创建的 Shapes 集合（包含所有拆分后的子项）。
    /// </summary>
    /// <returns>拆分后的子形状集合。</returns>
    [ReturnValueConvert]
    IExcelShapes? Ungroup();

    /// <summary>
    /// 将组合形状置于所有其他形状的顶层。
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将组合形状置于所有其他形状的底层。
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 删除此组合形状（及其所有子项）。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择此组合形状（激活并选中）。
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切当前组合形状到剪贴板。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制当前组合形状到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 创建当前组合形状的副本并放置在相同位置。
    /// </summary>
    void Duplicate();

    /// <summary>
    /// 将当前组合形状作为图片复制到剪贴板。
    /// </summary>
    /// <param name="Appearance">指定图片外观，可以是屏幕外观或打印外观。</param>
    /// <param name="Format">指定图片格式，可以是位图或矢量图格式。</param>
    void CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter,
                          XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);
}