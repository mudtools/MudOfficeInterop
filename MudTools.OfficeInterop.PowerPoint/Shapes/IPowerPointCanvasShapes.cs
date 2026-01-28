//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示画布上形状的集合，提供在画布上创建和操作形状的方法。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCanvasShapes : IEnumerable<IPowerPointShape?>, IDisposable
{
    /// <summary>
    /// 获取创建此形状集合的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取形状集合的父对象。
    /// </summary>
    /// <value>父对象，通常是画布或文档。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中形状的数量。
    /// </summary>
    /// <value>集合中形状的总数。</value>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的形状。
    /// </summary>
    /// <param name="index">要获取的形状的索引或名称。</param>
    /// <returns>指定索引处的形状对象。</returns>
    IPowerPointShape? this[int index] { get; }

    /// <summary>
    /// 通过索引获取集合中的形状。
    /// </summary>
    /// <param name="name">要获取的形状的索引或名称。</param>
    /// <returns>指定索引处的形状对象。</returns>
    IPowerPointShape? this[string name] { get; }

    /// <summary>
    /// 添加一个标注形状到画布。
    /// </summary>
    /// <param name="type">标注的类型。</param>
    /// <param name="left">标注左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">标注左上角相对于画布上边缘的位置（磅）。</param>
    /// <param name="width">标注的宽度（磅）。</param>
    /// <param name="height">标注的高度（磅）。</param>
    /// <returns>新添加的标注形状。</returns>
    IPowerPointShape? AddCallout([ComNamespace("MsCore")] MsoCalloutType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加一个连接符形状到画布。
    /// </summary>
    /// <param name="type">连接符的类型。</param>
    /// <param name="beginX">连接符起点的X坐标（磅）。</param>
    /// <param name="beginY">连接符起点的Y坐标（磅）。</param>
    /// <param name="endX">连接符终点的X坐标（磅）。</param>
    /// <param name="endY">连接符终点的Y坐标（磅）。</param>
    /// <returns>新添加的连接符形状。</returns>
    IPowerPointShape? AddConnector([ComNamespace("MsCore")] MsoConnectorType type, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加一条曲线形状到画布。
    /// </summary>
    /// <param name="safeArrayOfPoints">定义曲线顶点坐标的数组。</param>
    /// <returns>新添加的曲线形状。</returns>
    IPowerPointShape? AddCurve(object safeArrayOfPoints);

    /// <summary>
    /// 添加一个标签形状到画布。
    /// </summary>
    /// <param name="orientation">标签中文本的方向。</param>
    /// <param name="left">标签左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">标签左上角相对于画布上边缘的位置（磅）。</param>
    /// <param name="width">标签的宽度（磅）。</param>
    /// <param name="height">标签的高度（磅）。</param>
    /// <returns>新添加的标签形状。</returns>
    IPowerPointShape? AddLabel([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 添加一条线段到画布。
    /// </summary>
    /// <param name="beginX">线段起点的X坐标（磅）。</param>
    /// <param name="beginY">线段起点的Y坐标（磅）。</param>
    /// <param name="endX">线段终点的X坐标（磅）。</param>
    /// <param name="endY">线段终点的Y坐标（磅）。</param>
    /// <returns>新添加的线段形状。</returns>
    IPowerPointShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加一张图片到画布。
    /// </summary>
    /// <param name="fileName">图片文件的路径和名称。</param>
    /// <param name="linkToFile">指示是否将图片链接到文件。</param>
    /// <param name="saveWithDocument">指示是否将图片与文档一起保存。</param>
    /// <param name="left">图片左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">图片左上角相对于画布上边缘的位置（磅）。</param>
    /// <param name="width">图片的宽度（磅），-1表示使用原始宽度。</param>
    /// <param name="height">图片的高度（磅），-1表示使用原始高度。</param>
    /// <returns>新添加的图片形状。</returns>
    IPowerPointShape? AddPicture(string fileName, [ConvertTriState] bool linkToFile, [ConvertTriState] bool saveWithDocument, float left, float top, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加一条折线到画布。
    /// </summary>
    /// <param name="safeArrayOfPoints">定义折线顶点坐标的数组。</param>
    /// <returns>新添加的折线形状。</returns>
    IPowerPointShape? AddPolyline(object safeArrayOfPoints);

    /// <summary>
    /// 添加一个自动形状到画布。
    /// </summary>
    /// <param name="type">自动形状的类型。</param>
    /// <param name="left">形状左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">形状左上角相对于画布上边缘的位置（磅）。</param>
    /// <param name="width">形状的宽度（磅）。</param>
    /// <param name="height">形状的高度（磅）。</param>
    /// <returns>新添加的自动形状。</returns>
    IPowerPointShape? AddShape([ComNamespace("MsCore")] MsoAutoShapeType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加一个艺术字形状到画布。
    /// </summary>
    /// <param name="presetTextEffect">预设的文本效果。</param>
    /// <param name="text">要在艺术字中显示的文本。</param>
    /// <param name="fontName">要使用的字体名称。</param>
    /// <param name="fontSize">字体大小（磅）。</param>
    /// <param name="fontBold">指示文本是否为粗体。</param>
    /// <param name="fontItalic">指示文本是否为斜体。</param>
    /// <param name="left">艺术字左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">艺术字左上角相对于画布上边缘的位置（磅）。</param>
    /// <returns>新添加的艺术字形状。</returns>
    IPowerPointShape? AddTextEffect([ComNamespace("MsCore")] MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top);

    /// <summary>
    /// 添加一个文本框到画布。
    /// </summary>
    /// <param name="orientation">文本框中文本的方向。</param>
    /// <param name="left">文本框左上角相对于画布左边缘的位置（磅）。</param>
    /// <param name="top">文本框左上角相对于画布上边缘的位置（磅）。</param>
    /// <param name="width">文本框的宽度（磅）。</param>
    /// <param name="height">文本框的高度（磅）。</param>
    /// <returns>新添加的文本框形状。</returns>
    IPowerPointShape? AddTextbox([ComNamespace("MsCore")] MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 创建一个自由形状构建器，用于构建自定义的自由形状。
    /// </summary>
    /// <param name="editingType">编辑类型。</param>
    /// <param name="x1">自由形状起点的X坐标（磅）。</param>
    /// <param name="y1">自由形状起点的Y坐标（磅）。</param>
    /// <returns>用于构建自由形状的自由形状构建器对象。</returns>
    IPowerPointFreeformBuilder? BuildFreeform([ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1);

    /// <summary>
    /// 获取一个形状范围，包含指定的形状。
    /// </summary>
    /// <param name="index">形状的索引、名称或索引数组。</param>
    /// <returns>包含指定形状的形状范围对象。</returns>
    IPowerPointShapeRange? Range(object index);

    /// <summary>
    /// 选择画布上的所有形状。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取画布的背景形状。
    /// </summary>
    /// <value>表示画布背景的形状对象。</value>
    IPowerPointShape? Background { get; }
}