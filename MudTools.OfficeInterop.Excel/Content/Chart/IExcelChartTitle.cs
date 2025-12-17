//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ChartTitle 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartTitle 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelChartTitle : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取图表标题的父对象
    /// 对应 ChartTitle.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图表标题所在的 Application 对象
    /// 对应 ChartTitle.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    /// <summary>
    /// 获取或设置图表标题的名称
    /// 对应 ChartTitle.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置图表标题的文本内容
    /// 对应 ChartTitle.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置图表标题的说明文字
    /// </summary>
    string Caption { get; set; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置图表标题的左边距
    /// 对应 ChartTitle.Left 属性
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图表标题的顶边距
    /// 对应 ChartTitle.Top 属性
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取图表标题的宽度
    /// 对应 ChartTitle.Width 属性
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取图表标题的高度
    /// 对应 ChartTitle.Height 属性
    /// </summary>
    double Height { get; }

    #endregion

    #region 格式设置
    /// <summary>
    /// 获取图表标题的边框对象
    /// 对应 ChartTitle.Border 属性
    /// </summary>
    IExcelBorder? Border { get; }
    /// <summary>
    /// 获取图表标题的字体对象 
    /// 对应 ChartTitle.Font 属性
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取样式的内部格式对象
    /// 对应 Style.Interior 属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取图表标题的字符对象，用于对标题文本进行字符级格式设置
    /// 对应 ChartTitle.Characters 属性
    /// </summary>
    IExcelCharacters? Characters { get; }

    /// <summary>
    /// 获取或设置是否自动缩放字体
    /// 对应 ChartTitle.AutoScaleFont 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取图表标题的背景填充对象 
    /// 对应 ChartTitle.Format.Fill 或 ChartTitle.Interior 属性
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取或设置图表标题的水平对齐方式
    /// 对应 ChartTitle.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置图表标题的垂直对齐方式
    /// 对应 ChartTitle.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置图表标题的位置类型
    /// 对应 ChartTitle.Position 属性
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取或设置使用 R1C1 引用样式的本地化公式
    /// 对应 ChartTitle.FormulaR1C1Local 属性
    /// </summary>
    string FormulaR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置本地化公式
    /// 对应 ChartTitle.FormulaLocal 属性
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置使用 R1C1 引用样式的公式
    /// 对应 ChartTitle.FormulaR1C1 属性
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置图表标题的公式
    /// 对应 ChartTitle.Formula 属性
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置图表标题是否具有阴影效果
    /// 对应 ChartTitle.Shadow 属性
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置图表标题的阅读顺序
    /// 对应 ChartTitle.ReadingOrder 属性
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置图表标题的方向
    /// 对应 ChartTitle.Orientation 属性
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置图表标题是否包含在图表布局中
    /// 对应 ChartTitle.IncludeInLayout 属性
    /// </summary>
    bool IncludeInLayout { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 选择图表标题
    /// 对应 ChartTitle.Select 方法
    /// </summary>
    void Select();

    /// <summary>
    /// 删除图表标题
    /// 对应 ChartTitle.Delete 方法
    /// </summary>
    void Delete();
    #endregion
}
