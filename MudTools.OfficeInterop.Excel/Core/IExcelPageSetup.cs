//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Excel;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel PageSetup 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PageSetup 的安全访问和操作
/// </summary>
public interface IExcelPageSetup : IDisposable
{
    #region 页面设置

    /// <summary>
    /// 获取工作表中所有页面的集合
    /// </summary>
    IExcelPages? Pages { get; }

    /// <summary>
    /// 获取文档的偶数页设置
    /// </summary>
    IExcelPage? EvenPage { get; }

    /// <summary>
    /// 获取文档的首页设置
    /// </summary>
    IExcelPage? FirstPage { get; }


    /// <summary>
    /// 获取页面页眉中部的图片对象
    /// 对应 PageSetup.CenterHeaderPicture 属性
    /// </summary>
    IExcelGraphic? CenterHeaderPicture { get; }

    /// <summary>
    /// 获取页面页脚中部的图片对象
    /// 对应 PageSetup.CenterFooterPicture 属性
    /// </summary>
    IExcelGraphic CenterFooterPicture { get; }

    /// <summary>
    /// 获取页面页眉左侧的图片对象
    /// 对应 PageSetup.LeftHeaderPicture 属性
    /// </summary>
    IExcelGraphic LeftHeaderPicture { get; }

    /// <summary>
    /// 获取页面页脚左侧的图片对象
    /// 对应 PageSetup.LeftFooterPicture 属性
    /// </summary>
    IExcelGraphic LeftFooterPicture { get; }

    /// <summary>
    /// 获取页面页眉右侧的图片对象
    /// 对应 PageSetup.RightHeaderPicture 属性
    /// </summary>
    IExcelGraphic RightHeaderPicture { get; }

    /// <summary>
    /// 获取页面页脚右侧的图片对象
    /// 对应 PageSetup.RightFooterPicture 属性
    /// </summary>
    IExcelGraphic RightFooterPicture { get; }

    /// <summary>
    /// 获取或设置页面方向（纵向或横向）
    /// 对应 PageSetup.Orientation 属性
    /// </summary>
    XlPageOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置纸张大小
    /// 对应 PageSetup.PaperSize 属性
    /// </summary>
    XlPaperSize PaperSize { get; set; }

    /// <summary>
    /// 获取或设置页面缩放比例（10-400）
    /// 对应 PageSetup.Zoom 属性
    /// </summary>
    int Zoom { get; set; }

    /// <summary>
    /// 获取或设置是否适合页面宽度
    /// 对应 PageSetup.FitToPagesWide 属性
    /// </summary>
    int FitToPagesWide { get; set; }

    /// <summary>
    /// 获取或设置是否适合页面高度
    /// 对应 PageSetup.FitToPagesTall 属性
    /// </summary>
    int FitToPagesTall { get; set; }


    /// <summary>
    /// 获取或设置是否为黑白打印
    /// 对应 PageSetup.BlackAndWhite 属性
    /// </summary>
    bool BlackAndWhite { get; set; }

    /// <summary>
    /// 获取或设置是否为单色打印
    /// 对应 PageSetup.PrintComments 属性
    /// </summary>
    XlPrintLocation PrintComments { get; set; }

    /// <summary>
    /// 获取或设置打印错误处理方式
    /// 对应 PageSetup.PrintErrors 属性
    /// </summary>
    XlPrintErrors PrintErrors { get; set; }

    #endregion

    #region 页边距设置

    /// <summary>
    /// 获取或设置左边距（英寸）
    /// 对应 PageSetup.LeftMargin 属性
    /// </summary>
    double LeftMargin { get; set; }

    /// <summary>
    /// 获取或设置右边距（英寸）
    /// 对应 PageSetup.RightMargin 属性
    /// </summary>
    double RightMargin { get; set; }

    /// <summary>
    /// 获取或设置上边距（英寸）
    /// 对应 PageSetup.TopMargin 属性
    /// </summary>
    double TopMargin { get; set; }

    /// <summary>
    /// 获取或设置下边距（英寸）
    /// 对应 PageSetup.BottomMargin 属性
    /// </summary>
    double BottomMargin { get; set; }

    /// <summary>
    /// 获取或设置页眉边距（英寸）
    /// 对应 PageSetup.HeaderMargin 属性
    /// </summary>
    double HeaderMargin { get; set; }

    /// <summary>
    /// 获取或设置页脚边距（英寸）
    /// 对应 PageSetup.FooterMargin 属性
    /// </summary>
    double FooterMargin { get; set; }

    /// <summary>
    /// 获取或设置居中方式（水平居中）
    /// 对应 PageSetup.CenterHorizontally 属性
    /// </summary>
    bool CenterHorizontally { get; set; }

    /// <summary>
    /// 获取或设置居中方式（垂直居中）
    /// 对应 PageSetup.CenterVertically 属性
    /// </summary>
    bool CenterVertically { get; set; }

    #endregion

    #region 页眉页脚设置

    /// <summary>
    /// 获取或设置左页眉内容
    /// 对应 PageSetup.LeftHeader 属性
    /// </summary>
    string LeftHeader { get; set; }

    /// <summary>
    /// 获取或设置中页眉内容
    /// 对应 PageSetup.CenterHeader 属性
    /// </summary>
    string CenterHeader { get; set; }

    /// <summary>
    /// 获取或设置右页眉内容
    /// 对应 PageSetup.RightHeader 属性
    /// </summary>
    string RightHeader { get; set; }

    /// <summary>
    /// 获取或设置左页脚内容
    /// 对应 PageSetup.LeftFooter 属性
    /// </summary>
    string LeftFooter { get; set; }

    /// <summary>
    /// 获取或设置中页脚内容
    /// 对应 PageSetup.CenterFooter 属性
    /// </summary>
    string CenterFooter { get; set; }

    /// <summary>
    /// 获取或设置右页脚内容
    /// 对应 PageSetup.RightFooter 属性
    /// </summary>
    string RightFooter { get; set; }

    #endregion

    #region 打印选项
    /// <summary>
    /// 获取或设置一个值，该值指示是否为奇数页和偶数页使用不同的页眉和页脚
    /// 对应 PageSetup.OddAndEvenPagesHeaderFooter 属性
    /// </summary>
    bool OddAndEvenPagesHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页眉和页脚是否随文档缩放
    /// 对应 PageSetup.ScaleWithDocHeaderFooter 属性
    /// </summary>
    bool ScaleWithDocHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页眉和页脚是否与页边距对齐
    /// 对应 PageSetup.AlignMarginsHeaderFooter 属性
    /// </summary>
    bool AlignMarginsHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置打印顺序
    /// 对应 PageSetup.Order 属性
    /// </summary>
    XlOrder Order { get; set; }

    /// <summary>
    /// 获取或设置是否以草稿模式打印
    /// 对应 PageSetup.Draft 属性
    /// </summary>
    bool Draft { get; set; }
    /// <summary>
    /// 获取或设置打印质量
    /// </summary>
    object? PrintQuality { get; set; }

    /// <summary>
    /// 获取或设置是否打印网格线
    /// 对应 PageSetup.PrintGridlines 属性
    /// </summary>
    bool PrintGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否打印行列标题
    /// 对应 PageSetup.PrintHeadings 属性
    /// </summary>
    bool PrintHeadings { get; set; }

    /// <summary>
    /// 获取或设置是否打印注释
    /// 对应 PageSetup.PrintNotes 属性
    /// </summary>
    bool PrintNotes { get; set; }

    /// <summary>
    /// 获取或设置是否打印标题行
    /// 对应 PageSetup.PrintTitleRows 属性
    /// </summary>
    string PrintTitleRows { get; set; }

    /// <summary>
    /// 获取或设置是否打印标题列
    /// 对应 PageSetup.PrintTitleColumns 属性
    /// </summary>
    string PrintTitleColumns { get; set; }

    /// <summary>
    /// 获取或设置打印区域
    /// 对应 PageSetup.PrintArea 属性
    /// </summary>
    string PrintArea { get; set; }

    /// <summary>
    /// 获取或设置是否从第一页开始编号
    /// 对应 PageSetup.FirstPageNumber 属性
    /// </summary>
    int FirstPageNumber { get; set; }

    /// <summary>
    /// 获取或设置是否不同奇偶页页眉页脚
    /// 对应 PageSetup.DifferentFirstPageHeaderFooter 属性
    /// </summary>
    bool DifferentFirstPageHeaderFooter { get; set; }

    #endregion

    #region 页面编号和日期

    /// <summary>
    /// 获取或设置是否显示页码
    /// </summary>
    bool ShowPageNumbers { get; set; }

    /// <summary>
    /// 获取或设置是否显示日期
    /// </summary>
    bool ShowDate { get; set; }

    /// <summary>
    /// 获取或设置是否显示时间
    /// </summary>
    bool ShowTime { get; set; }

    /// <summary>
    /// 获取或设置是否显示文件名
    /// </summary>
    bool ShowFileName { get; set; }

    /// <summary>
    /// 获取或设置是否显示工作表名
    /// </summary>
    bool ShowSheetName { get; set; }

    /// <summary>
    /// 获取或设置是否显示路径
    /// </summary>
    bool ShowPath { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 应用页面设置
    /// </summary>
    void Apply();

    /// <summary>
    /// 重置页面设置为默认值
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制页面设置
    /// </summary>
    /// <param name="source">源页面设置对象</param>
    void Copy(IExcelPageSetup source);

    /// <summary>
    /// 获取标准页眉页脚代码
    /// </summary>
    /// <param name="type">页眉页脚类型</param>
    /// <returns>标准代码</returns>
    string GetStandardHeaderFooterCode(int type);

    /// <summary>
    /// 设置自定义页眉页脚
    /// </summary>
    /// <param name="section">区域（左、中、右）</param>
    /// <param name="position">位置（页眉、页脚）</param>
    /// <param name="text">文本内容</param>
    void SetCustomHeaderFooter(int section, int position, string text);

    #endregion
}