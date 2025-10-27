//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Shapes 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Shapes 的安全访问和操作
/// </summary>
public interface IExcelShapes : IEnumerable<IExcelShape>, IDisposable
{
    /// <summary>
    /// 获取形状集合中的形状数量
    /// 对应 Shapes.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的形状对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    IExcelShape? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    IExcelShape? this[string name] { get; }

    /// <summary>
    /// 添加文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddTextbox(int orientation, float left, float top, float width, float height);

    /// <summary>
    /// 添加矩形形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddRectangle(float left, float top, float width, float height);

    /// <summary>
    /// 添加椭圆形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddEllipse(float left, float top, float width, float height);

    /// <summary>
    /// 添加图表形状(增强版)
    /// </summary>
    /// <param name="style">图表样式索引，-1表示默认样式</param>
    /// <param name="chartType">图表类型</param>
    /// <param name="left">图表左边距，null表示默认值</param>
    /// <param name="top">图表顶边距，null表示默认值</param>
    /// <param name="width">图表宽度，null表示默认值</param>
    /// <param name="height">图表高度，null表示默认值</param>
    /// <param name="newLayout">是否使用新布局，null表示默认设置</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddChart2(int style = -1, MsoChartType chartType = MsoChartType.xlPie,
                     float? left = null, float? top = null, float?
                     width = null, float? height = null, bool? newLayout = null);

    /// <summary>
    /// 添加线条形状
    /// </summary>
    /// <param name="x1">起点X坐标</param>
    /// <param name="y1">起点Y坐标</param>
    /// <param name="x2">终点X坐标</param>
    /// <param name="y2">终点Y坐标</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddLine(float x1, float y1, float x2, float y2);

    /// <summary>
    /// 添加图片形状
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddPicture(string filename, bool linkToFile, bool saveWithDocument, float left, float top, float width, float height);

    /// <summary>
    /// 添加图表形状
    /// </summary>
    /// <param name="Type">图表类型</param>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddDiagram(MsoDiagramType Type, float Left, float Top, float Width, float Height);

    /// <summary>
    /// 添加画布形状
    /// </summary>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddCanvas(float Left, float Top, float Width, float Height);

    /// <summary>
    /// 添加图表形状
    /// </summary>
    /// <param name="XlChartType">图表类型</param>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddChart(MsoChartType XlChartType, float Left, float Top, float Width, float Height);

    /// <summary>
    /// 添加SmartArt形状
    /// </summary>
    /// <param name="Layout">SmartArt布局</param>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddSmartArt(IOfficeSmartArtLayout Layout, float Left, float Top, float Width, float Height);

    /// <summary>
    /// 添加多段线形状
    /// </summary>
    /// <param name="points">点坐标数组</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddPolyline(float[,] points);

    /// <summary>
    /// 添加曲线形状
    /// </summary>
    /// <param name="points">点坐标数组</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddCurve(float[,] points);

    /// <summary>
    /// 添加标签形状
    /// </summary>
    /// <param name="Orientation">文本方向</param>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <param name="Width">宽度</param>
    /// <param name="Height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddLabel(MsoTextOrientation Orientation, float Left, float Top, float Width, float Height);

    /// <summary>
    /// 添加连接符形状
    /// </summary>
    /// <param name="type">连接符类型</param>
    /// <param name="BeginX">起点X坐标</param>
    /// <param name="BeginY">起点Y坐标</param>
    /// <param name="EndX">终点X坐标</param>
    /// <param name="EndY">终点Y坐标</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddConnector(MsoConnectorType type, float BeginX, float BeginY, float EndX, float EndY);

    /// <summary>
    /// 添加自定义形状
    /// </summary>
    /// <param name="shapeType">形状类型</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddShape(MsoAutoShapeType shapeType, float left, float top, float width, float height);

    /// <summary>
    /// 添加艺术字形状
    /// </summary>
    /// <param name="PresetTextEffect">预设文本效果样式</param>
    /// <param name="Text">艺术字文本内容</param>
    /// <param name="FontName">字体名称</param>
    /// <param name="FontSize">字体大小</param>
    /// <param name="FontBold">是否粗体</param>
    /// <param name="FontItalic">是否斜体</param>
    /// <param name="Left">左边距</param>
    /// <param name="Top">顶边距</param>
    /// <returns>新创建的形状对象</returns>
    public IExcelShape? AddTextEffect(
       MsoPresetTextEffect PresetTextEffect,
       string Text, string FontName,
       float FontSize, bool FontBold,
       bool FontItalic, float Left, float Top);

    /// <summary>
    /// 获取指定索引或名称的形状区域对象
    /// </summary>
    /// <param name="index">形状的索引或名称</param>
    /// <returns>形状区域对象</returns>
    IExcelShapeRange? Range(string index);

    /// <summary>
    /// 选择所有形状
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 删除所有形状
    /// </summary>
    void DeleteAll();
}