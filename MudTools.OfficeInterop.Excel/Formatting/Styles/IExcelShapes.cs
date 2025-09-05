//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
    IExcelShape this[int index] { get; }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    IExcelShape this[string name] { get; }

    /// <summary>
    /// 添加文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape AddTextbox(int orientation, double left, double top, double width, double height);

    /// <summary>
    /// 添加矩形形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape AddRectangle(double left, double top, double width, double height);

    /// <summary>
    /// 添加椭圆形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape AddEllipse(double left, double top, double width, double height);

    /// <summary>
    /// 添加线条形状
    /// </summary>
    /// <param name="x1">起点X坐标</param>
    /// <param name="y1">起点Y坐标</param>
    /// <param name="x2">终点X坐标</param>
    /// <param name="y2">终点Y坐标</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape AddLine(double x1, double y1, double x2, double y2);

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
    IExcelShape? AddPicture(string filename, bool linkToFile, bool saveWithDocument, double left, double top, double width, double height);


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