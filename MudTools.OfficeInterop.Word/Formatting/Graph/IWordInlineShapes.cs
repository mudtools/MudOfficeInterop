//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中内嵌形状集合的封装接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordInlineShapes : IEnumerable<IWordInlineShape?>, IOfficeObject<IWordInlineShapes, MsWord.InlineShapes>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取内嵌形状的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的内嵌形状对象
    /// </summary>
    /// <param name="index">内嵌形状索引（从1开始）</param>
    /// <returns>内嵌形状对象</returns>
    IWordInlineShape? this[int index] { get; }

    /// <summary>
    /// 插入一个由边框环绕的空白的 1 英寸见方的 Microsoft Word 图片对象。
    /// </summary>
    /// <param name="range">必需。Range 对象。新图形的位置。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? New(IWordRange range);

    /// <summary>
    /// 添加图片内嵌形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <returns>新添加的内嵌形状对象</returns>
    IWordInlineShape? AddPicture(string fileName, bool linkToFile = false, bool saveWithDocument = true);

    /// <summary>
    /// 添加OLE对象内嵌形状
    /// </summary>
    /// <param name="classType">OLE对象类类型</param>
    /// <param name="fileName">文件名</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="displayAsIcon">是否以图标显示</param>
    /// <param name="iconFileName">图标文件路径</param>
    /// <param name="iconIndex">图标索引</param>
    /// <param name="iconLabel">图标标签</param>
    /// <returns>新创建的内嵌形状对象</returns>
    IWordInlineShape? AddOLEObject(string? classType = null, string? fileName = null, bool? linkToFile = false,
                                 bool? displayAsIcon = false, string? iconFileName = null, int? iconIndex = 0,
                                 string? iconLabel = null);


    /// <summary>
    /// 基于图像文件向当前文档添加水平线。
    /// </summary>
    /// <param name="fileName">必需。要用作水平线的图像的文件名。</param>
    /// <param name="range">可选项。Microsoft Word 放置水平线之上的范围。如果省略此参数，Word 将在当前选定内容之上放置水平线。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddHorizontalLine(string fileName, IWordRange? range = null);

    /// <summary>
    /// 向当前文档添加水平线。
    /// </summary>
    /// <param name="range">可选项。Microsoft Word 放置水平线之上的范围。如果省略此参数，Word 将在当前选定内容之上放置水平线。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddHorizontalLineStandard(IWordRange? range = null);

    /// <summary>
    /// 基于图像文件向当前文档添加图片项目符号。
    /// </summary>
    /// <param name="fileName">必需。要用作图片项目符号的图像的文件名。</param>
    /// <param name="range">可选项。Microsoft Word 添加图片项目符号的范围。Word 将图片项目符号添加到范围内的每个段落。如果省略此参数，Word 将图片项目符号添加到当前选定内容中的每个段落。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddPictureBullet(string fileName, IWordRange? range = null);

    /// <summary>
    /// 将指定类型的图表作为内嵌形状插入活动文档，并打开 Microsoft Office Excel，其中包含 Microsoft Office Word 用于创建图表的默认数据的工作表。
    /// </summary>
    /// <param name="type">指定要创建的图表类型。</param>
    /// <param name="range">指定图表绑定到的文本。如果指定了 Range，图表将放置在范围中第一个段落的开头。如果省略此参数，则自动选择范围，图表相对于页面的顶部和左侧边缘定位。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddChart([ComNamespace("MsCore")] XlChartType type = XlChartType.xlArea, IWordRange? range = null);

    /// <summary>
    /// 添加具有指定样式和类型的图表。
    /// </summary>
    /// <param name="style">图表样式。</param>
    /// <param name="type">图表类型。</param>
    /// <param name="range">放置图表的范围。</param>
    /// <param name="newLayout">新布局。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddChart2(int style = -1, [ComNamespace("MsCore")] XlChartType type = XlChartType.xlArea, IWordRange? range = null, object? newLayout = null);

    /// <summary>
    /// 将 SmartArt 图形作为内嵌形状插入活动文档。
    /// </summary>
    /// <param name="layout">指定 SmartArt 图形布局的 SmartArtLayout 对象。</param>
    /// <param name="range">指定 SmartArt 图形绑定到的文本。如果指定了 Range，SmartArt 图形将放置在范围中第一个段落的开头。如果省略此参数，则自动选择范围，SmartArt 图形相对于页面的顶部和左侧边缘定位。</param>
    /// <returns>要插入的 SmartArt 图形。</returns>
    IWordInlineShape? AddSmartArt(IOfficeSmartArtLayout layout, IWordRange? range = null);

    /// <summary>
    /// 添加 Web 视频作为内嵌形状。
    /// </summary>
    /// <param name="embedCode">嵌入代码。</param>
    /// <param name="videoWidth">视频宽度。</param>
    /// <param name="videoHeight">视频高度。</param>
    /// <param name="posterFrameImage">海报帧图像。</param>
    /// <param name="url">视频 URL。</param>
    /// <param name="range">放置视频的范围。</param>
    /// <returns>新创建的 InlineShape 对象。</returns>
    IWordInlineShape? AddWebVideo(string embedCode, int? videoWidth, int? videoHeight, string? posterFrameImage = null, string? url = null, IWordRange? range = null);

}