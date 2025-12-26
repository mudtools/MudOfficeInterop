//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示一个形状集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeCanvasShapes : IOfficeObject<IOfficeCanvasShapes>, IEnumerable<IOfficeShape?>, IDisposable
{
    /// <summary>
    /// 通过索引获取形状
    /// </summary>
    /// <param name="index">形状的索引（从1开始）</param>
    /// <returns>指定索引处的形状</returns>
    IOfficeShape? this[int index] { get; }

    /// <summary>
    /// 通过名称获取形状
    /// </summary>
    /// <param name="name">形状的名称</param>
    /// <returns>具有指定名称的形状</returns>
    IOfficeShape? this[string name] { get; }

    /// <summary>
    /// 获取当前形状集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前形状集合中形状的数量
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 添加一个标注形状
    /// </summary>
    /// <param name="type">标注类型</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <param name="width">形状宽度</param>
    /// <param name="height">形状高度</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddCallout(MsoCalloutType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加一个连接符形状
    /// </summary>
    /// <param name="type">连接符类型</param>
    /// <param name="beginX">起始点X坐标</param>
    /// <param name="beginY">起始点Y坐标</param>
    /// <param name="endX">终点X坐标</param>
    /// <param name="endY">终点Y坐标</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddConnector(MsoConnectorType type, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加一个曲线形状
    /// </summary>
    /// <param name="safeArrayOfPoints">安全数组格式的点坐标</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddCurve(object safeArrayOfPoints);

    /// <summary>
    /// 添加一个标签形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <param name="width">形状宽度</param>
    /// <param name="height">形状高度</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddLabel(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 添加一条直线形状
    /// </summary>
    /// <param name="beginX">起始点X坐标</param>
    /// <param name="beginY">起始点Y坐标</param>
    /// <param name="endX">终点X坐标</param>
    /// <param name="endY">终点Y坐标</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 添加一个图片形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <param name="width">形状宽度，默认为-1</param>
    /// <param name="height">形状高度，默认为-1</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddPicture(string fileName, [ConvertTriState] bool linkToFile, [ConvertTriState] bool saveWithDocument, float left, float top, float width = -1f, float height = -1f);

    /// <summary>
    /// 添加一个多段线形状
    /// </summary>
    /// <param name="safeArrayOfPoints">安全数组格式的点坐标</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddPolyline(object safeArrayOfPoints);

    /// <summary>
    /// 添加一个自选形状
    /// </summary>
    /// <param name="type">自选形状类型</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <param name="width">形状宽度</param>
    /// <param name="height">形状高度</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddShape(MsoAutoShapeType type, float left, float top, float width, float height);

    /// <summary>
    /// 添加一个艺术字形状
    /// </summary>
    /// <param name="presetTextEffect">预设文本效果</param>
    /// <param name="text">艺术字文本</param>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="fontBold">是否粗体</param>
    /// <param name="fontItalic">是否斜体</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddTextEffect(MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top);

    /// <summary>
    /// 添加一个文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">形状左侧位置</param>
    /// <param name="top">形状顶部位置</param>
    /// <param name="width">形状宽度</param>
    /// <param name="height">形状高度</param>
    /// <returns>新添加的形状对象</returns>
    IOfficeShape? AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 获取形状范围
    /// </summary>
    /// <param name="index">索引</param>
    /// <returns>形状范围对象</returns>
    IOfficeShapeRange? Range(object index);

}
