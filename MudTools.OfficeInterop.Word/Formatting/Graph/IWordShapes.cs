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
public interface IWordShapes : IEnumerable<IWordShape>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取形状集合中的形状数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取形状（从1开始）。
    /// </summary>
    IWordShape this[int index] { get; }

    /// <summary>
    /// 根据名称获取形状。
    /// </summary>
    IWordShape this[string name] { get; }

    /// <summary>
    /// 添加 SmartArt 图形。
    /// </summary>
    /// <param name="Layout">SmartArt 布局类型。</param>
    /// <param name="Left">形状左边距。</param>
    /// <param name="Top">形状上边距。</param>
    /// <param name="Width">形状宽度。</param>
    /// <param name="Height">形状高度。</param>
    /// <param name="Anchor">锚点范围，决定形状在文档中的位置。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape AddSmartArt(IOfficeSmartArtLayout Layout, float Left, float Top, float Width, float Height, IWordRange? Anchor = null);

    /// <summary>
    /// 添加自定义形状。
    /// </summary>
    /// <param name="Type">自定义形状的类型。</param>
    /// <param name="Left">形状左边距。</param>
    /// <param name="Top">形状上边距。</param>
    /// <param name="Width">形状宽度。</param>
    /// <param name="Height">形状高度。</param>
    /// <param name="Anchor">锚点范围，决定形状在文档中的位置。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape AddShape(MsoAutoShapeType Type, float Left, float Top, float Width, float Height, IWordRange? Anchor = null);


    /// <summary>
    /// 添加文本框形状。
    /// </summary>
    /// <param name="orientation">文本方向。</param>
    /// <param name="left">左边距。</param>
    /// <param name="top">上边距。</param>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape? AddTextbox(MsoTextOrientation orientation, float left, float top,
        float width, float height);

    /// <summary>
    /// 添加矩形形状。
    /// </summary>
    /// <param name="left">左边距。</param>
    /// <param name="top">上边距。</param>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape? AddRectangle(float left, float top, float width, float height);

    /// <summary>
    /// 添加线条形状。
    /// </summary>
    /// <param name="beginX">起始点X坐标。</param>
    /// <param name="beginY">起始点Y坐标。</param>
    /// <param name="endX">结束点X坐标。</param>
    /// <param name="endY">结束点Y坐标。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加图片形状。
    /// </summary>
    /// <param name="fileName">图片文件路径。</param>
    /// <param name="linkToFile">是否链接到文件。</param>
    /// <param name="saveWithDocument">是否与文档一起保存。</param>
    /// <param name="left">左边距。</param>
    /// <param name="top">上边距。</param>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape? AddPicture(string fileName, bool linkToFile, bool saveWithDocument,
        float left, float top, float width, float height);

    /// <summary>
    /// 添加图表形状。
    /// </summary>
    /// <param name="type">图表类型。</param>
    /// <param name="left">左边距。</param>
    /// <param name="top">上边距。</param>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    /// <returns>新添加的形状。</returns>
    IWordShape? AddChart(MsoChartType type, float left, float top, float width, float height);

    IWordShape? AddOLEObject(string? classType = null, string? fileName = null, bool? linkToFile = false,
       bool? displayAsIcon = false, string? iconFileName = null, int? iconIndex = null, string? iconLabel = null,
       float? left = null, float? top = null, float? width = null, float? height = null, object? anchor = null);


    /// <summary>
    /// 检查形状是否存在。
    /// </summary>
    /// <param name="name">形状名称。</param>
    /// <returns>是否存在。</returns>
    bool Contains(string name);

    /// <summary>
    /// 获取所有形状名称。
    /// </summary>
    /// <returns>形状名称列表。</returns>
    List<string> GetAllShapeNames();

    /// <summary>
    /// 根据形状类型获取形状名称列表。
    /// </summary>
    /// <param name="shapeType">形状类型。</param>
    /// <returns>形状名称列表。</returns>
    List<string> GetShapeNamesByType(MsoShapeType shapeType);

    /// <summary>
    /// 删除指定名称的形状。
    /// </summary>
    /// <param name="name">形状名称。</param>
    /// <returns>是否删除成功。</returns>
    bool DeleteShape(string name);

    /// <summary>
    /// 删除所有形状。
    /// </summary>
    void DeleteAll();

    /// <summary>
    /// 选择所有形状。
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取指定范围内的形状。
    /// </summary>
    /// <param name="range">范围。</param>
    /// <returns>形状集合。</returns>
    IWordShapeRange GetShapesInRange(IWordRange range);

    /// <summary>
    /// 获取指定类型的形状数量。
    /// </summary>
    /// <param name="shapeType">形状类型。</param>
    /// <returns>形状数量。</returns>
    int GetCountByType(MsoShapeType shapeType);
}