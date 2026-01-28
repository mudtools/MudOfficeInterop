//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示一组 Office 形状对象的集合接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
[ItemIndex]
public interface IOfficeShapes : IOfficeObject<IOfficeShapes, MsCore.Shapes>, IEnumerable<IOfficeShape?>, IDisposable
{
    /// <summary>
    /// 获取集合中的形状数量
    /// </summary>
    int Count { get; }


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
    /// 向集合中添加新形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的形状</returns>
    IOfficeShape? AddShape(MsoAutoShapeType type, float left, float top, float width, float height);

    /// <summary>
    /// 向集合中添加文本框
    /// </summary>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="orientation"></param>
    /// <returns>新添加的文本框形状</returns>
    IOfficeShape? AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 向集合中添加艺术字形状
    /// </summary>
    /// <param name="presetTextEffect">预设的艺术字效果</param>
    /// <param name="text">艺术字文本内容</param>
    /// <param name="fontName">字体名称</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="fontBold">是否粗体</param>
    /// <param name="fontItalic">是否斜体</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <returns>新添加的艺术字形状</returns>
    IOfficeShape? AddTextEffect(MsoPresetTextEffect presetTextEffect, string text, string fontName, float fontSize, [ConvertTriState] bool fontBold, [ConvertTriState] bool fontItalic, float left, float top);


    /// <summary>
    /// 向集合中添加标注形状
    /// </summary>
    /// <param name="type">标注类型</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的标注形状</returns>
    IOfficeShape? AddCallout(MsoCalloutType type, float left, float top, float width, float height);

    /// <summary>
    /// 向集合中添加连接符形状
    /// </summary>
    /// <param name="type">连接符类型</param>
    /// <param name="beginX">起点X坐标</param>
    /// <param name="beginY">起点Y坐标</param>
    /// <param name="endX">终点X坐标</param>
    /// <param name="endY">终点Y坐标</param>
    /// <returns>新添加的连接符形状</returns>
    IOfficeShape? AddConnector(MsoConnectorType type, float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 向集合中添加曲线形状
    /// </summary>
    /// <param name="safeArrayOfPoints">点坐标数组</param>
    /// <returns>新添加的曲线形状</returns>
    IOfficeShape? AddCurve(object safeArrayOfPoints);

    /// <summary>
    /// 向集合中添加标签形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的标签形状</returns>
    IOfficeShape? AddLabel(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 向集合中添加直线形状
    /// </summary>
    /// <param name="beginX">起点X坐标</param>
    /// <param name="beginY">起点Y坐标</param>
    /// <param name="endX">终点X坐标</param>
    /// <param name="endY">终点Y坐标</param>
    /// <returns>新添加的直线形状</returns>
    IOfficeShape? AddLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 向集合中添加图片形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否随文档保存</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度，默认为-1表示自动调整</param>
    /// <param name="height">高度，默认为-1表示自动调整</param>
    /// <returns>新添加的图片形状</returns>
    IOfficeShape? AddPicture(string fileName, [ConvertTriState] bool linkToFile, [ConvertTriState] bool saveWithDocument, float left, float top, float width = -1f, float height = -1f);

    /// <summary>
    /// 向集合中添加多段线形状
    /// </summary>
    /// <param name="safeArrayOfPoints">点坐标数组</param>
    /// <returns>新添加的多段线形状</returns>
    IOfficeShape? AddPolyline(object safeArrayOfPoints);

    /// <summary>
    /// 获取指定索引的形状范围
    /// </summary>
    /// <param name="index">形状索引</param>
    /// <returns>指定索引的形状范围</returns>
    IOfficeShapeRange? Range(int index);
}