//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表中的绘图对象的接口。
/// 绘图对象包括图表、形状、图片、文本框等可视元素。
/// 该接口继承自IDisposable，确保正确释放资源。
/// </summary>
public interface IExcelDrawing : IExcelComGraphObject, IDisposable
{
    /// <summary>
    /// 获取父级绘图对象集合
    /// </summary>
    IExcelDrawingObjects ParentDrawing { get; }

    /// <summary>
    /// 获取关联的工作表
    /// </summary>
    IExcelWorksheet Worksheet { get; }

    /// <summary>
    /// 获取或设置绘图对象的文本内容（如果是文本框或形状）
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取绘图对象的字体属性（如果是文本对象）
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取或设置绘图对象的水平对齐方式
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置绘图对象的垂直对齐方式
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }



    /// <summary>
    /// 调整绘图对象大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    void Resize(double width, double height);

}