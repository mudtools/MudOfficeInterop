//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 幻灯片中的形状集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShapes : IOfficeObject<IPowerPointShapes, MsPowerPoint.Shapes>, IEnumerable<IPowerPointShape?>, IDisposable
{
    /// <summary>
    /// 获取创建此形状集合的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此形状集合的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此形状集合的父对象。
    /// </summary>
    /// <value>表示此形状集合父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取形状集合中的形状数量。
    /// </summary>
    /// <value>集合中的形状数量。</value>
    int Count { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定形状。
    /// </summary>
    /// <param name="index">要获取的形状的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? this[int index] { get; }

    /// <summary>
    /// 通过索引或名称获取集合中的指定形状。
    /// </summary>
    /// <param name="name">要获取的形状的索引（从1开始）或名称。</param>
    /// <value>指定索引或名称对应的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? this[string name] { get; }

    /// <summary>
    /// 添加标注形状到集合中。
    /// </summary>
    /// <param name="type">标注类型。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddCallout([ComNamespace("MsCore")] MsoCalloutType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加连接线形状到集合中。
    /// </summary>
    /// <param name="type">连接线类型。</param>
    /// <param name="beginX">起始点的 X 坐标（以磅为单位）。</param>
    /// <param name="beginY">起始点的 Y 坐标（以磅为单位）。</param>
    /// <param name="endX">结束点的 X 坐标（以磅为单位）。</param>
    /// <param name="endY">结束点的 Y 坐标（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddConnector([ComNamespace("MsCore")] MsoConnectorType type, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加曲线形状到集合中。
    /// </summary>
    /// <param name="safeArrayOfPoints">包含曲线点的数组。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddCurve(object safeArrayOfPoints);

    /// <summary>
    /// 添加标签形状到集合中。
    /// </summary>
    /// <param name="orientation">文本方向。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddLabel([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 添加直线形状到集合中。
    /// </summary>
    /// <param name="beginX">起始点的 X 坐标（以磅为单位）。</param>
    /// <param name="beginY">起始点的 Y 坐标（以磅为单位）。</param>
    /// <param name="endX">结束点的 X 坐标（以磅为单位）。</param>
    /// <param name="endY">结束点的 Y 坐标（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加图片形状到集合中。
    /// </summary>
    /// <param name="fileName">图片文件名称。</param>
    /// <param name="linkToFile">指示是否链接到文件的布尔值。</param>
    /// <param name="saveWithDocument">指示是否随文档保存的布尔值。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用原始宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用原始高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddPicture(string fileName, [ConvertTriState] bool linkToFile, [ConvertTriState] bool saveWithDocument, float left, float top, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加折线形状到集合中。
    /// </summary>
    /// <param name="safeArrayOfPoints">包含折线点的数组。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddPolyline(object safeArrayOfPoints);

    /// <summary>
    /// 添加自选图形到集合中。
    /// </summary>
    /// <param name="type">自选图形类型。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddShape([ComNamespace("MsCore")] MsoAutoShapeType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加艺术字形状到集合中。
    /// </summary>
    /// <param name="presetTextEffect">预设的艺术字效果。</param>
    /// <param name="text">艺术字文本。</param>
    /// <param name="fontName">字体名称。</param>
    /// <param name="fontSize">字体大小。</param>
    /// <param name="fontBold">指示字体是否为粗体的布尔值。</param>
    /// <param name="fontItalic">指示字体是否为斜体的布尔值。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddTextEffect([ComNamespace("MsCore")] MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top);

    /// <summary>
    /// 添加文本框形状到集合中。
    /// </summary>
    /// <param name="orientation">文本方向。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddTextbox([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 创建自由形状构建器。
    /// </summary>
    /// <param name="editingType">编辑类型。</param>
    /// <param name="x1">第一个点的 X 坐标（以磅为单位）。</param>
    /// <param name="y1">第一个点的 Y 坐标（以磅为单位）。</param>
    /// <returns>新创建的 <see cref="IPowerPointFreeformBuilder"/> 对象。</returns>
    IPowerPointFreeformBuilder? BuildFreeform([ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1);

    /// <summary>
    /// 选择集合中的所有形状。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取指定形状的范围。
    /// </summary>
    /// <param name="index">要获取范围的形状的索引（从1开始）、索引数组或名称。</param>
    /// <returns>指定形状范围对应的 <see cref="IPowerPointShapeRange"/> 对象。</returns>
    IPowerPointShapeRange? Range([Optional] object index);

    /// <summary>
    /// 获取一个值，指示幻灯片是否有标题。
    /// </summary>
    /// <value>指示幻灯片是否有标题的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTitle { get; }

    /// <summary>
    /// 添加标题形状到集合中。
    /// </summary>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddTitle();

    /// <summary>
    /// 获取幻灯片的标题形状。
    /// </summary>
    /// <value>表示标题形状的 <see cref="IPowerPointShape"/> 对象。</value>
    IPowerPointShape? Title { get; }

    /// <summary>
    /// 获取幻灯片中的占位符集合。
    /// </summary>
    /// <value>表示占位符集合的 <see cref="IPowerPointPlaceholders"/> 对象。</value>
    IPowerPointPlaceholders? Placeholders { get; }

    /// <summary>
    /// 添加 OLE 对象形状到集合中。
    /// </summary>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <param name="className">对象的类名。</param>
    /// <param name="fileName">对象的文件名称。</param>
    /// <param name="displayAsIcon">指示是否显示为图标的布尔值。</param>
    /// <param name="iconFileName">图标文件的名称。</param>
    /// <param name="iconIndex">图标的索引。</param>
    /// <param name="iconLabel">图标的标签。</param>
    /// <param name="link">指示是否链接到文件的布尔值。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddOLEObject(float left = 0f, float top = 0f, float width = -1f, float height = -1f, string className = "", string fileName = "", [ConvertTriState] bool displayAsIcon = false, string iconFileName = "", int iconIndex = 0, string iconLabel = "", [ConvertTriState] bool link = false);

    /// <summary>
    /// 添加注释形状到集合中。
    /// </summary>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddComment(float left = 1f, float top = 1f, float width = 145f, float height = 145f);

    /// <summary>
    /// 添加占位符形状到集合中。
    /// </summary>
    /// <param name="type">占位符类型。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddPlaceholder(PpPlaceholderType type, float left = -1f, float top = -1f, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加媒体对象形状到集合中。
    /// </summary>
    /// <param name="fileName">媒体文件的名称。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddMediaObject(string fileName, float left = 0f, float top = 0f, float width = -1f, float height = -1f);

    /// <summary>
    /// 粘贴剪贴板内容为形状。
    /// </summary>
    /// <returns>粘贴的形状范围。</returns>
    IPowerPointShapeRange? Paste();

    /// <summary>
    /// 添加表格形状到集合中。
    /// </summary>
    /// <param name="numRows">表格的行数。</param>
    /// <param name="numColumns">表格的列数。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddTable(int numRows, int numColumns, float left = -1f, float top = -1f, float width = -1f, float height = -1f);

    /// <summary>
    /// 以特殊格式粘贴剪贴板内容为形状。
    /// </summary>
    /// <param name="dataType">粘贴的数据类型。</param>
    /// <param name="displayAsIcon">指示是否显示为图标的布尔值。</param>
    /// <param name="iconFileName">图标文件的名称。</param>
    /// <param name="iconIndex">图标的索引。</param>
    /// <param name="iconLabel">图标的标签。</param>
    /// <param name="link">指示是否链接到文件的布尔值。</param>
    /// <returns>粘贴的形状范围。</returns>
    IPowerPointShapeRange? PasteSpecial(PpPasteDataType dataType = PpPasteDataType.ppPasteDefault, [ConvertTriState] bool displayAsIcon = false, string iconFileName = "", int iconIndex = 0, string iconLabel = "", [ConvertTriState] bool link = false);

    /// <summary>
    /// 添加图表形状到集合中。
    /// </summary>
    /// <param name="type"></param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。</param>
    /// <param name="height">形状的高度（以磅为单位）。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddChart([ComNamespace("MsCore")] XlChartType type = XlChartType.xlArea, float left = -1, float top = -1f, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加媒体对象形状到集合中（增强版本）。
    /// </summary>
    /// <param name="fileName">媒体文件的名称。</param>
    /// <param name="linkToFile">指示是否链接到文件的布尔值。</param>
    /// <param name="saveWithDocument">指示是否随文档保存的布尔值。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddMediaObject2(string fileName, [ConvertTriState] bool linkToFile = false, [ConvertTriState] bool saveWithDocument = true, float left = 0f, float top = 0f, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加智能艺术形状到集合中。
    /// </summary>
    /// <param name="layout">智能艺术布局。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。值为-1表示使用默认位置。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddSmartArt(IOfficeSmartArtLayout layout, float left = -1f, float top = -1f, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加图片形状到集合中（增强版本）。
    /// </summary>
    /// <param name="fileName">图片文件名称。</param>
    /// <param name="linkToFile">指示是否链接到文件的布尔值。</param>
    /// <param name="saveWithDocument">指示是否随文档保存的布尔值。</param>
    /// <param name="left">形状的左边缘位置（以磅为单位）。</param>
    /// <param name="top">形状的上边缘位置（以磅为单位）。</param>
    /// <param name="width">形状的宽度（以磅为单位）。值为-1表示使用默认宽度。</param>
    /// <param name="height">形状的高度（以磅为单位）。值为-1表示使用默认高度。</param>
    /// <param name="compress">图片压缩选项。</param>
    /// <returns>新添加的 <see cref="IPowerPointShape"/> 对象。</returns>
    IPowerPointShape? AddPicture2(string fileName, [ConvertTriState] bool linkToFile, [ConvertTriState] bool saveWithDocument, float left, float top, float width = -1f, float height = -1f, [ComNamespace("MsCore")] MsoPictureCompress compress = MsoPictureCompress.msoPictureCompressDocDefault);
}