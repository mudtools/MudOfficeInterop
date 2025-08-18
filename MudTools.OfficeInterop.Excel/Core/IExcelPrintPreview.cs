//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PrintPreview 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PrintPreview 的安全访问和操作
/// </summary>
public interface IExcelPrintPreview : IDisposable
{
    #region 基础属性   

    /// <summary>
    /// 获取打印预览窗口的父对象（通常是工作表或工作簿）
    /// </summary>
    object Parent { get; }

    #endregion

    #region 显示设置

    /// <summary>
    /// 获取或设置打印预览的缩放比例（10-400）
    /// </summary>
    int Zoom { get; set; }

    /// <summary>
    /// 获取或设置是否显示页眉
    /// </summary>
    bool ShowHeaders { get; set; }

    /// <summary>
    /// 获取或设置是否显示页脚
    /// </summary>
    bool ShowFooters { get; set; }

    /// <summary>
    /// 获取或设置是否显示网格线
    /// </summary>
    bool ShowGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否显示行列标题
    /// </summary>
    bool ShowHeadings { get; set; }

    /// <summary>
    /// 获取或设置是否显示注释
    /// </summary>
    int ShowComments { get; set; }
    #endregion

    #region 页面设置

    /// <summary>
    /// 获取或设置页面方向（纵向或横向）
    /// </summary>
    int Orientation { get; set; }

    /// <summary>
    /// 获取或设置纸张大小
    /// </summary>
    int PaperSize { get; set; }
    #endregion

    #region 页边距设置

    /// <summary>
    /// 获取或设置左边距（英寸）
    /// </summary>
    double LeftMargin { get; set; }

    /// <summary>
    /// 获取或设置右边距（英寸）
    /// </summary>
    double RightMargin { get; set; }

    /// <summary>
    /// 获取或设置上边距（英寸）
    /// </summary>
    double TopMargin { get; set; }

    /// <summary>
    /// 获取或设置下边距（英寸）
    /// </summary>
    double BottomMargin { get; set; }

    /// <summary>
    /// 获取或设置页眉边距（英寸）
    /// </summary>
    double HeaderMargin { get; set; }

    /// <summary>
    /// 获取或设置页脚边距（英寸）
    /// </summary>
    double FooterMargin { get; set; }

    #endregion

    #region 页眉页脚设置

    /// <summary>
    /// 获取或设置左页眉内容
    /// </summary>
    string LeftHeader { get; set; }

    /// <summary>
    /// 获取或设置中页眉内容
    /// </summary>
    string CenterHeader { get; set; }

    /// <summary>
    /// 获取或设置右页眉内容
    /// </summary>
    string RightHeader { get; set; }

    /// <summary>
    /// 获取或设置左页脚内容
    /// </summary>
    string LeftFooter { get; set; }

    /// <summary>
    /// 获取或设置中页脚内容
    /// </summary>
    string CenterFooter { get; set; }

    /// <summary>
    /// 获取或设置右页脚内容
    /// </summary>
    string RightFooter { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 显示打印预览窗口
    /// </summary>
    /// <param name="enableChanges">是否允许在预览中进行更改</param>
    void Show(bool enableChanges = true);

    /// <summary>
    /// 刷新打印预览显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 打印当前预览的内容
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="collate">是否逐份打印</param>
    void Print(int copies = 1, bool collate = true);

    /// <summary>
    /// 导出预览为PDF文件
    /// </summary>
    /// <param name="filename">PDF文件路径</param>
    void ExportToPDF(string filename);

    #endregion

    #region 高级功能

    /// <summary>
    /// 获取或设置是否显示黑白预览
    /// </summary>
    bool BlackAndWhite { get; set; }

    /// <summary>
    /// 获取或设置是否显示注释预览
    /// </summary>
    bool PrintNotes { get; set; }

    /// <summary>
    /// 获取或设置是否显示网格线预览
    /// </summary>
    bool PrintGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否显示行列标题预览
    /// </summary>
    bool PrintHeadings { get; set; }

    /// <summary>
    /// 获取或设置打印区域预览
    /// </summary>
    string PrintArea { get; set; }

    /// <summary>
    /// 获取或设置打印标题行预览
    /// </summary>
    string PrintTitleRows { get; set; }

    /// <summary>
    /// 获取或设置打印标题列预览
    /// </summary>
    string PrintTitleColumns { get; set; }

    #endregion
}