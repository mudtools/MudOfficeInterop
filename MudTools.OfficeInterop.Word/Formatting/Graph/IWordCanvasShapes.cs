//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示绘图画布中的形状集合。
/// 此接口提供对画布中形状的管理功能，包括添加、选择和操作各种类型的形状。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordCanvasShapes : IEnumerable<IWordShape?>, IOfficeObject<IWordCanvasShapes, MsWord.CanvasShapes>, IDisposable
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
    /// 获取集合中的形状数量。
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 通过索引获取指定的形状。
    /// </summary>
    /// <param name="index">形状的序号位置或表示形状名称的字符串。</param>
    /// <returns>指定索引处的形状对象。</returns>
    IWordShape? this[int index] { get; }

    /// <summary>
    /// 通过索引获取指定的形状。
    /// </summary>
    /// <param name="name">形状的序号位置或表示形状名称的字符串。</param>
    /// <returns>指定索引处的形状对象。</returns>
    IWordShape? this[string name] { get; }

    /// <summary>
    /// 向绘图画布添加无边框线标注。
    /// </summary>
    /// <param name="type">标注的类型。</param>
    /// <param name="left">标注边框左边缘的位置（以磅为单位）。</param>
    /// <param name="top">标注边框上边缘的位置（以磅为单位）。</param>
    /// <param name="width">标注边框的宽度（以磅为单位）。</param>
    /// <param name="height">标注边框的高度（以磅为单位）。</param>
    /// <returns>新创建的标注形状对象。</returns>
    IWordShape? AddCallout([ComNamespace("MsCore")] MsoCalloutType type, float left, float top, float width, float height);

    /// <summary>
    /// 在绘图画布中两个形状之间添加连接线。
    /// </summary>
    /// <param name="type">连接器的类型。</param>
    /// <param name="beginX">标记连接器起点的水平位置。</param>
    /// <param name="beginY">标记连接器起点的垂直位置。</param>
    /// <param name="endX">标记连接器终点的水平位置。</param>
    /// <param name="endY">标记连接器终点的垂直位置。</param>
    /// <returns>新创建的连接器形状对象。</returns>
    IWordShape? AddConnector([ComNamespace("MsCore")] MsoConnectorType type, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 向绘图画布添加贝塞尔曲线。
    /// </summary>
    /// <param name="safeArrayOfPoints">指定曲线顶点和控制点的坐标对数组。
    /// 指定的第一个点是起始顶点，接下来的两个点是第一个贝塞尔段的控制点。
    /// 然后，对于曲线的每个附加段，指定一个顶点和两个控制点。
    /// 指定的最后一个点是曲线的结束顶点。
    /// 注意：必须始终指定 3n + 1 个点，其中 n 是曲线中的段数。</param>
    /// <returns>新创建的曲线形状对象。</returns>
    IWordShape? AddCurve(ref object safeArrayOfPoints);

    /// <summary>
    /// 向绘图画布添加文本标签。
    /// </summary>
    /// <param name="orientation">文本的方向。</param>
    /// <param name="left">标签左边缘相对于绘图画布左边缘的位置（以磅为单位）。</param>
    /// <param name="top">标签上边缘相对于绘图画布上边缘的位置（以磅为单位）。</param>
    /// <param name="width">标签的宽度（以磅为单位）。</param>
    /// <param name="height">标签的高度（以磅为单位）。</param>
    /// <returns>新创建的标签形状对象。</returns>
    IWordShape? AddLabel([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 向绘图画布添加线条。
    /// </summary>
    /// <param name="beginX">线条起点相对于绘图画布的水平位置（以磅为单位）。</param>
    /// <param name="beginY">线条起点相对于绘图画布的垂直位置（以磅为单位）。</param>
    /// <param name="endX">线条终点相对于绘图画布的水平位置（以磅为单位）。</param>
    /// <param name="endY">线条终点相对于绘图画布的垂直位置（以磅为单位）。</param>
    /// <returns>新创建的线条形状对象。</returns>
    IWordShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 向绘图画布添加图片。
    /// </summary>
    /// <param name="fileName">图片的路径和文件名。</param>
    /// <param name="linkToFile">指示是否将图片链接到创建它的文件。
    /// False 表示图片应为文件的独立副本。默认值为 False。</param>
    /// <param name="saveWithDocument">指示是否将链接的图片与文档一起保存。默认值为 False。</param>
    /// <param name="left">新图片左边缘相对于绘图画布的位置（以磅为单位）。</param>
    /// <param name="top">新图片上边缘相对于绘图画布的位置（以磅为单位）。</param>
    /// <param name="width">图片的宽度（以磅为单位）。</param>
    /// <param name="height">图片的高度（以磅为单位）。</param>
    /// <returns>新创建的图片形状对象。</returns>
    IWordShape? AddPicture(string fileName, bool? linkToFile = null, bool? saveWithDocument = null, float? left = null, float? top = null, float? width = null, float? height = null);

    /// <summary>
    /// 向绘图画布添加开放或闭合的多边形。
    /// </summary>
    /// <param name="safeArrayOfPoints">指定折线顶点坐标对的数组。</param>
    /// <returns>新创建的多边形形状对象。</returns>
    IWordShape? AddPolyline(object safeArrayOfPoints);

    /// <summary>
    /// 向绘图画布添加自选图形。
    /// </summary>
    /// <param name="type">要返回的形状类型。可以是任何 MsoAutoShapeType 常量。</param>
    /// <param name="left">自选图形左边缘的位置（以磅为单位）。</param>
    /// <param name="top">自选图形上边缘的位置（以磅为单位）。</param>
    /// <param name="width">自选图形的宽度（以磅为单位）。</param>
    /// <param name="height">自选图形的高度（以磅为单位）。</param>
    /// <returns>新创建的自选图形对象。</returns>
    IWordShape? AddShape(int type, float left, float top, float width, float height);

    /// <summary>
    /// 向绘图画布添加艺术字形状。
    /// </summary>
    /// <param name="presetTextEffect">预设的文本效果。MsoPresetTextEffect 常量的值对应于"艺术字库"对话框中列出的格式（从左到右，从上到下编号）。</param>
    /// <param name="text">艺术字中的文本。</param>
    /// <param name="fontName">艺术字中使用的字体名称。</param>
    /// <param name="fontSize">艺术字中使用的字体大小（以磅为单位）。</param>
    /// <param name="fontBold">指示是否将艺术字字体加粗的值。</param>
    /// <param name="fontItalic">指示是否将艺术字字体倾斜的值。</param>
    /// <param name="left">艺术字形状左边缘相对于绘图画布左边缘的位置（以磅为单位）。</param>
    /// <param name="top">艺术字形状上边缘相对于绘图画布上边缘的位置（以磅为单位）。</param>
    /// <returns>新创建的艺术字形状对象。</returns>
    IWordShape? AddTextEffect([ComNamespace("MsCore")] MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top);

    /// <summary>
    /// 向绘图画布添加文本框。
    /// </summary>
    /// <param name="orientation">文本的方向。某些常量可能不可用，具体取决于您选择或安装的语言支持（例如美式英语）。</param>
    /// <param name="left">文本框左边缘的位置（以磅为单位）。</param>
    /// <param name="top">文本框上边缘的位置（以磅为单位）。</param>
    /// <param name="width">文本框的宽度（以磅为单位）。</param>
    /// <param name="height">文本框的高度（以磅为单位）。</param>
    /// <returns>新创建的文本框形状对象。</returns>
    IWordShape? AddTextbox([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 构建自由曲线对象。
    /// </summary>
    /// <param name="editingType">第一个节点的编辑属性。</param>
    /// <param name="x1">自由曲线中第一个节点相对于文档左上角的 X 位置。</param>
    /// <param name="y1">自由曲线中第一个节点相对于文档左上角的 Y 位置。</param>
    /// <returns>新创建的自由曲线构建器对象。</returns>
    IWordFreeformBuilder? BuildFreeform([ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1);

    /// <summary>
    /// 指定要包含在特定范围内的形状。
    /// </summary>
    /// <param name="index">指定哪些形状应包含在指定范围内。
    /// 可以是整数，指定形状在 Shapes 集合中的索引号；
    /// 可以是字符串，指定形状的名称；
    /// 也可以是包含整数或字符串的数组。</param>
    /// <returns>指定范围内的形状范围对象。</returns>
    IWordShapeRange? Range(object index);

    /// <summary>
    /// 指定要包含在特定范围内的形状。
    /// </summary>
    /// <param name="index">指定哪些形状应包含在指定范围内。
    /// 可以是整数，指定形状在 Shapes 集合中的索引号；</param>
    /// <returns>指定范围内的形状范围对象。</returns>
    IWordShapeRange? Range(int index);

    /// <summary>
    /// 指定要包含在特定范围内的形状。
    /// </summary>
    /// <param name="name">指定哪些形状应包含在指定范围内。
    /// 可以是字符串，指定形状的名称；
    /// 也可以是包含整数或字符串的数组。</param>
    /// <returns>指定范围内的形状范围对象。</returns>
    IWordShapeRange? Range(string name);

    /// <summary>
    /// 选择文档主故事、画布或页眉页脚中的所有形状。
    /// </summary>
    void SelectAll();
}