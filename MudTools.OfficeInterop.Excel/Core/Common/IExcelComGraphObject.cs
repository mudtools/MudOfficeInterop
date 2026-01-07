//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 公共图形对象。
/// </summary>
public interface IExcelComGraphObject : IDisposable
{
    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取绘图对象的边框格式设置
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取绘图对象的内部格式设置
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取图表对象的左上角单元格
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取绘图对象的范围（如果是图表或列表对象）
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取图表对象所在的父对象（通常是工作簿）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置图表对象的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取图表对象的索引位置
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置打印工作表时是否打印该控件
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 获取或设置绘图对象是否锁定
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置图表是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置图表对象的左边距
    /// 对应 ChartObject.Left 属性
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图表对象的顶边距
    /// 对应 ChartObject.Top 属性
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置图表对象的宽度
    /// 对应 ChartObject.Width 属性
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置图表对象的高度
    /// 对应 ChartObject.Height 属性
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 复制对象
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切图表对象
    /// 对应 ChartObject.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除图表对象
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择绘图对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 移动绘图对象到指定位置
    /// </summary>
    /// <param name="left">新左侧位置</param>
    /// <param name="top">新顶部位置</param>
    void Move(double left, double top);

    /// <summary>
    /// 将图表对象置于最前面
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将图表对象置于最后面
    /// </summary>
    void SendToBack();
}
