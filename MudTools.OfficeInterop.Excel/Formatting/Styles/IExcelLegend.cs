//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Legend 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Legend 的安全访问和操作
/// </summary>
public interface IExcelLegend : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置图例的名称
    /// 对应 Legend.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取图例的父对象
    /// 对应 Legend.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取图例所在的 Application 对象
    /// 对应 Legend.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置图例的左边距
    /// 对应 Legend.Left 属性
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图例的顶边距
    /// 对应 Legend.Top 属性
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置图例的宽度
    /// 对应 Legend.Width 属性
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置图例的高度
    /// 对应 Legend.Height 属性
    /// </summary>
    double Height { get; set; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取图例的字体对象
    /// 对应 Legend.Font 属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取或设置是否自动缩放字体
    /// 对应 Legend.AutoScaleFont 属性
    /// </summary>
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取图例的背景填充对象
    /// 对应 Legend.Format.Fill 或 Legend.Interior 属性
    /// </summary>
    IExcelChartFillFormat Fill { get; }


    /// <summary>
    /// 获取绘图区的边框对象
    /// </summary>
    IExcelBorder Border { get; }

    /// <summary>
    /// 获取样式的内部格式对象
    /// 对应 Style.Interior 属性
    /// </summary>
    IExcelInterior Interior { get; }

    /// <summary>
    /// 获取或设置图例的位置
    /// 对应 Legend.Position 属性
    /// </summary>
    XlLegendPosition Position { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 选择图例
    /// 对应 Legend.Select 方法
    /// </summary>
    void Select();

    /// <summary>
    /// 删除图例
    /// 对应 Legend.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除图例内容
    /// 对应 Legend.Clear 方法
    /// </summary>
    void Clear();

    #endregion
}
