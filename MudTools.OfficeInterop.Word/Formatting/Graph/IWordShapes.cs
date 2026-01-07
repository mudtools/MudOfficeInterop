//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shapes 的接口，用于操作文档中的形状集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordShapes : IEnumerable<IWordShape?>, IOfficeObject<IWordShapes, MsWord.Shapes>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取形状集合中的形状数量。
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
    /// 向文档添加无边框线标注。
    /// 返回表示标注的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="type">标注的类型。
    /// msoCalloutOne：沿标注框左边缘垂直向下放置标注线。
    /// msoCalloutTwo：沿标注框左边缘斜向下放置标注线。
    /// msoCalloutThree：沿标注框左边缘水平放置，然后斜向下放置标注线。
    /// msoCalloutFour：沿标注文本框左边缘放置标注线。
    /// msoCalloutMixed：表示范围或选择中存在多个 MsoCalloutType 的返回值。</param>
    /// <param name="left">标注边框左边缘的位置（以磅为单位）。</param>
    /// <param name="top">标注边框上边缘的位置（以磅为单位）。</param>
    /// <param name="width">标注边框的宽度（以磅为单位）。</param>
    /// <param name="height">标注边框的高度（以磅为单位）。</param>
    /// <param name="anchor">表示标注绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，标注相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的标注形状对象。</returns>
    IWordShape? AddCallout([ComNamespace("MsCore")] MsoCalloutType type, float left, float top, float width, float height, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加贝塞尔曲线。返回表示曲线的 Shape 对象。
    /// </summary>
    /// <param name="safeArrayOfPoints">指定曲线顶点和控制点的坐标对数组。
    /// 指定的第一个点是起始顶点，接下来的两个点是第一个贝塞尔段的控制点。
    /// 然后，对于曲线的每个附加段，指定一个顶点和两个控制点。
    /// 指定的最后一个点是曲线的结束顶点。
    /// 注意：必须始终指定 3n + 1 个点，其中 n 是曲线中的段数。</param>
    /// <param name="anchor">表示曲线绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，曲线相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的曲线形状对象。</returns>
    IWordShape? AddCurve(object safeArrayOfPoints, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加文本标签。返回表示文本标签的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="orientation">文本的方向。
    /// 某些常量可能不可用，具体取决于您选择或安装的语言支持（例如美式英语）。</param>
    /// <param name="left">标签左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">标签上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">标签的宽度（以磅为单位）。</param>
    /// <param name="height">标签的高度（以磅为单位）。</param>
    /// <param name="anchor">表示标签绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，标签相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的标签形状对象。</returns>
    IWordShape? AddLabel([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加线条。返回表示线条的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="beginX">线条起点相对于定位点的水平位置（以磅为单位）。</param>
    /// <param name="beginY">线条起点相对于定位点的垂直位置（以磅为单位）。</param>
    /// <param name="endX">线条终点相对于定位点的水平位置（以磅为单位）。</param>
    /// <param name="endY">线条终点相对于定位点的垂直位置（以磅为单位）。</param>
    /// <param name="anchor">表示线条绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，线条相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的线条形状对象。</returns>
    IWordShape? AddLine(float beginX, float beginY, float endX, float endY, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加图片。返回表示图片的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="fileName">图片的路径和文件名。</param>
    /// <param name="linkToFile">True 将图片链接到创建它的文件，False 使图片成为文件的独立副本。默认值为 False。</param>
    /// <param name="saveWithDocument">True 将链接的图片与文档一起保存。默认值为 False。</param>
    /// <param name="left">新图片左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">新图片上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">图片的宽度（以磅为单位）。</param>
    /// <param name="height">图片的高度（以磅为单位）。</param>
    /// <param name="anchor">表示图片绑定文本的 Range 对象。
    /// <para> 如果指定了 anchor，则定位点位于定位范围第一段的开头。</para>
    /// <para>如果省略此参数，则自动选择定位范围，图片相对于页面的顶部和左边缘定位。</para></param>
    /// <returns>新创建的图片形状对象。</returns>
    IWordShape? AddPicture(string fileName, bool? linkToFile = null, bool? saveWithDocument = null, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加开放或闭合的多边形。返回表示多边形的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="safeArrayOfPoints">指定折线顶点坐标对的数组。</param>
    /// <param name="anchor">表示折线绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，折线相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的多边形形状对象。</returns>
    IWordShape? AddPolyline(object safeArrayOfPoints, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加自选图形。返回表示自选图形的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="type">要返回的形状类型。可以是任何 MsoAutoShapeType 常量。</param>
    /// <param name="left">自选图形左边缘的位置（以磅为单位）。</param>
    /// <param name="top">自选图形上边缘的位置（以磅为单位）。</param>
    /// <param name="width">自选图形的宽度（以磅为单位）。</param>
    /// <param name="height">自选图形的高度（以磅为单位）。</param>
    /// <param name="anchor">表示自选图形绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，自选图形相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的自选图形对象。</returns>
    IWordShape? AddShape(int type, float left, float top, float width, float height, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加艺术字形状。返回表示艺术字的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="presetTextEffect">预设的文本效果。MsoPresetTextEffect 常量的值对应于"艺术字库"对话框中列出的格式（从左到右，从上到下编号）。</param>
    /// <param name="text">艺术字中的文本。</param>
    /// <param name="fontName">艺术字中使用的字体名称。</param>
    /// <param name="fontSize">艺术字中使用的字体大小（以磅为单位）。</param>
    /// <param name="fontBold">指示是否将艺术字字体加粗的值。</param>
    /// <param name="fontItalic">指示是否将艺术字字体倾斜的值。</param>
    /// <param name="left">艺术字形状左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">艺术字形状上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="anchor">表示艺术字绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，艺术字相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的艺术字形状对象。</returns>
    IWordShape? AddTextEffect([ComNamespace("MsCore")] MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加文本框。返回表示文本框的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="orientation">文本的方向。某些常量可能不可用，具体取决于您选择或安装的语言支持（例如美式英语）。</param>
    /// <param name="left">文本框左边缘的位置（以磅为单位）。</param>
    /// <param name="top">文本框上边缘的位置（以磅为单位）。</param>
    /// <param name="width">文本框的宽度（以磅为单位）。</param>
    /// <param name="height">文本框的高度（以磅为单位）。</param>
    /// <param name="anchor">表示文本框绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，文本框相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的文本框形状对象。</returns>
    IWordShape? AddTextbox([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height, IWordRange? anchor = null);

    /// <summary>
    /// 构建自由曲线对象。返回表示正在构建的自由曲线的 FreeformBuilder 对象。
    /// </summary>
    /// <param name="editingType">第一个节点的编辑属性。可以是 msoEditingAuto 或 msoEditingCorner，不能是 msoEditingSmooth 或 msoEditingSymmetric。</param>
    /// <param name="x1">自由曲线中第一个节点相对于文档左上角的位置（以磅为单位）。</param>
    /// <param name="y1">自由曲线中第一个节点相对于文档左上角的位置（以磅为单位）。</param>
    /// <returns>新创建的自由曲线构建器对象。</returns>
    IWordFreeformBuilder? BuildFreeform([ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1);

    /// <summary>
    /// 返回 ShapeRange 对象。
    /// </summary>
    /// <param name="index">指定要包含在指定范围内的形状。
    /// 可以是整数，指定形状在 Shapes 集合中的索引号；
    /// 可以是字符串，指定形状的名称；
    /// 也可以是包含整数或字符串的对象数组。</param>
    /// <returns>指定范围内的形状范围对象。</returns>
    IWordShapeRange? Range(int index);

    /// <summary>
    /// 返回 ShapeRange 对象。
    /// </summary>
    /// <param name="name">指定要包含在指定范围内的形状。
    /// 可以是整数，指定形状在 Shapes 集合中的索引号；
    /// 可以是字符串，指定形状的名称；
    /// 也可以是包含整数或字符串的对象数组。</param>
    /// <returns>指定范围内的形状范围对象。</returns>
    IWordShapeRange? Range(string name);

    /// <summary>
    /// 选择文档主故事、画布或页眉页脚中的所有形状。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 创建 OLE 对象。返回表示新 OLE 对象的 Shape 对象。
    /// </summary>
    /// <param name="classType">用于激活指定 OLE 对象的应用程序名称。</param>
    /// <param name="fileName">创建对象的文件。如果省略此参数，则使用当前文件夹。
    /// 必须为对象指定 ClassType 或 FileName 参数之一，但不能同时指定两者。</param>
    /// <param name="linkToFile">True 将 OLE 对象链接到创建它的文件，False 使 OLE 对象成为文件的独立副本。
    /// 如果为 ClassType 指定了值，则 LinkToFile 参数必须为 False。默认值为 False。</param>
    /// <param name="displayAsIcon">True 将 OLE 对象显示为图标。默认值为 False。</param>
    /// <param name="iconFileName">包含要显示的图标的文件。</param>
    /// <param name="iconIndex">IconFileName 中图标的索引号。
    /// 指定文件中的图标顺序对应于选中"显示为图标"复选框时"更改图标"对话框（"插入"菜单，"对象"对话框）中图标的显示顺序。
    /// 文件中的第一个图标的索引号为 0。如果 IconFileName 中不存在具有给定索引号的图标，则使用索引号 1 的图标（文件中的第二个图标）。
    /// 默认值为 0。</param>
    /// <param name="iconLabel">图标下方显示的标签（标题）。</param>
    /// <param name="left">新对象左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">新对象上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">OLE 对象的宽度（以磅为单位）。</param>
    /// <param name="height">OLE 对象的高度（以磅为单位）。</param>
    /// <param name="anchor">表示 OLE 对象绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果未指定 anchor，则自动放置定位点，OLE 对象相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的 OLE 对象形状。</returns>
    IWordShape? AddOLEObject(object? classType = null, string? fileName = null, bool? linkToFile = null, bool? displayAsIcon = null, string? iconFileName = null, int? iconIndex = null, string? iconLabel = null, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null);

    /// <summary>
    /// 创建 ActiveX 控件（以前称为 OLE 控件）。返回表示新 ActiveX 控件的 Shape 对象。
    /// </summary>
    /// <param name="classType">要创建的 ActiveX 控件的编程标识符。</param>
    /// <param name="left">新对象左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">新对象上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">ActiveX 控件的宽度（以磅为单位）。</param>
    /// <param name="height">ActiveX 控件的高度（以磅为单位）。</param>
    /// <param name="anchor">表示 ActiveX 控件绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动放置定位点，ActiveX 控件相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的 ActiveX 控件形状。</returns>
    IWordShape? AddOLEControl(object? classType = null, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null);

    /// <summary>
    /// 向文档添加绘图画布。返回表示绘图画布的 Shape 对象并将其添加到 Shapes 集合。
    /// </summary>
    /// <param name="left">绘图画布左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">绘图画布上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">绘图画布的宽度（以磅为单位）。</param>
    /// <param name="height">绘图画布的高度（以磅为单位）。</param>
    /// <param name="anchor">表示画布绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，画布相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的绘图画布形状对象。</returns>
    IWordShape? AddCanvas(float left, float top, float width, float height, IWordRange? anchor = null);

    /// <summary>
    /// 将指定类型的图表作为形状插入活动文档，并打开 Microsoft Office Excel，其中包含 Microsoft Office Word 用于创建图表的默认数据表。
    /// </summary>
    /// <param name="type">指定要创建的图表类型。</param>
    /// <param name="left">图表左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">图表上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">图表的宽度（以磅为单位）。</param>
    /// <param name="height">图表的高度（以磅为单位）。</param>
    /// <param name="anchor">表示图表绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，图表相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的图表形状对象。</returns>
    IWordShape? AddChart([ComNamespace("MsCore")] XlChartType type = XlChartType.xl3DLine, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null);

    /// <summary>
    /// 将指定的 SmartArt 图形插入活动文档。
    /// </summary>
    /// <param name="layout">指定 SmartArt 图形布局的 SmartArtLayout 对象。</param>
    /// <param name="left">SmartArt 图形左边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="top">SmartArt 图形上边缘相对于定位点的位置（以磅为单位）。</param>
    /// <param name="width">SmartArt 图形的宽度。</param>
    /// <param name="height">SmartArt 图形的高度。</param>
    /// <param name="anchor">表示 SmartArt 图形绑定文本的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，SmartArt 图形相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的 SmartArt 图形形状对象。</returns>
    IWordShape? AddSmartArt(IOfficeSmartArtLayout layout, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null);

    /// <summary>
    /// 添加具有指定样式和类型的图表。
    /// </summary>
    /// <param name="style">图表的样式。</param>
    /// <param name="type">图表的类型。</param>
    /// <param name="left">图表左边缘的位置（以磅为单位）。</param>
    /// <param name="top">图表上边缘的位置（以磅为单位）。</param>
    /// <param name="width">图表的宽度（以磅为单位）。</param>
    /// <param name="height">图表的高度（以磅为单位）。</param>
    /// <param name="anchor">图表绑定的文本范围。</param>
    /// <param name="newLayout">新布局设置。</param>
    /// <returns>新创建的图表形状对象。</returns>
    IWordShape? AddChart2(int style = -1, [ComNamespace("MsCore")] XlChartType type = XlChartType.xl3DLine, float? left = null, float? top = null, float? width = null, float? height = null, IWordRange? anchor = null, object? newLayout = null);
}