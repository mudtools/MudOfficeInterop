//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel CellFormat 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.CellFormat 的安全访问和操作
/// </summary>
public interface IExcelCellFormat : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取单元格格式对象的父对象（通常是 Application）
    /// 对应 CellFormat.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取单元格格式对象所在的Application对象
    /// 对应 CellFormat.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置单元格的数字格式
    /// 对应 CellFormat.NumberFormat 属性
    /// </summary>
    object NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置单元格的水平对齐方式
    /// 对应 CellFormat.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置单元格的垂直对齐方式
    /// 对应 CellFormat.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置单元格的缩进量
    /// 对应 CellFormat.IndentLevel 属性
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取或设置单元格的方向
    /// 对应 CellFormat.Orientation 属性
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置单元格是否自动缩放字体
    /// 对应 CellFormat.ShrinkToFit 属性
    /// </summary>
    bool ShrinkToFit { get; set; }

    /// <summary>
    /// 获取或设置单元格是否自动换行
    /// 对应 CellFormat.WrapText 属性
    /// </summary>
    bool WrapText { get; set; }

    /// <summary>
    /// 获取或设置单元格的合并单元格属性
    /// 对应 CellFormat.MergeCells 属性
    /// </summary>
    bool MergeCells { get; set; }

    /// <summary>
    /// 获取或设置单元格的锁定状态
    /// 对应 CellFormat.Locked 属性
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置单元格的公式隐藏状态
    /// 对应 CellFormat.FormulaHidden 属性
    /// </summary>
    bool FormulaHidden { get; set; }

    string NumberFormatLocal { get; set; }
    #endregion

    #region 格式设置 (子对象)
    /// <summary>
    /// 获取单元格的字体对象
    /// 对应 CellFormat.Font 属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取单元格的背景填充对象
    /// 对应 CellFormat.Interior 属性
    /// </summary>
    IExcelInterior Interior { get; }

    /// <summary>
    /// 获取单元格的边框对象
    /// 对应 CellFormat.Borders 属性
    /// </summary>
    IExcelBorders Borders { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 清除单元格格式
    /// 对应 CellFormat.Clear 方法
    /// </summary>
    void Clear();
    #endregion
}
