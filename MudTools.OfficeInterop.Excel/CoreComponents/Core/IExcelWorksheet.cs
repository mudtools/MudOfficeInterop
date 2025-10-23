//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Worksheet 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Worksheet 的安全访问和操作
/// </summary>
public interface IExcelWorksheet : IExcelComSheet, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取工作表的命名区域集合
    /// 对应 Worksheet.Names 属性
    /// </summary>
    IExcelNames? Names { get; }

    IExcelPrintPreview? PrintPreview { get; }

    /// <summary>
    /// 获取工作表的垂直分页符集合
    /// </summary>
    IExcelVPageBreaks? VPageBreaks { get; }

    /// <summary>
    /// 获取工作表的水平分页符集合
    /// </summary>
    IExcelHPageBreaks? HPageBreaks { get; }

    /// <summary>
    /// 获取一个值，该值指示工作表中的图形对象是否受到保护
    /// </summary>
    bool ProtectDrawingObjects { get; }

    /// <summary>
    /// 获取一个值，该值指示工作表中的方案是否受到保护
    /// </summary>
    bool ProtectScenarios { get; }

    /// <summary>
    /// 获取或设置工作表标签颜色
    /// </summary>
    Color TabColor { get; set; }

    /// <summary>
    /// 获取工作表单元格
    /// </summary>
    IExcelCells? Cells { get; }

    /// <summary>
    /// 获取工作表中的循环引用单元格集合
    /// </summary>
    IExcelCells? CircularReference { get; }

    /// <summary>
    /// 获取工作表的排序操作对象
    /// </summary>
    IExcelSort? Sort { get; }

    /// <summary>
    /// 获取或设置标准列宽
    /// </summary>
    double StandardWidth { get; set; }

    /// <summary>
    /// 获取大纲（分级显示）设置对象
    /// </summary>
    IExcelOutline? Outline { get; }

    /// <summary>
    /// 获取或设置自动筛选模式状态
    /// </summary>
    bool AutoFilterMode { get; set; }

    /// <summary>
    /// 获取工作表的自动筛选器对象
    /// 对应 Worksheet.AutoFilter 属性
    /// </summary>
    IExcelAutoFilter? AutoFilter { get; }

    /// <summary>
    /// 获取或设置是否显示分页符
    /// </summary>
    bool DisplayPageBreaks { get; set; }


    /// <summary>
    /// 获取工作表当前是否处于筛选模式
    /// </summary>
    bool FilterMode { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否使用 Lotus 1-2-3 表达式计算规则
    /// </summary>
    bool TransitionExpEval { get; set; }


    /// <summary>
    /// 获取或设置一个值，该值指示是否启用大纲（分级显示）功能
    /// </summary>
    bool EnableOutlining { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否启用数据透视表功能
    /// </summary>
    bool EnablePivotTable { get; set; }

    /// <summary>
    /// 获取或设置工作表中可以选择的单元格类型
    /// </summary>
    XlEnableSelection EnableSelection { get; set; }

    /// <summary>
    /// 获取或设置在重新计算工作表时运行的宏名称
    /// </summary>
    string OnCalculate { get; set; }

    /// <summary>
    /// 获取或设置在工作表接收数据时运行的宏名称
    /// </summary>
    string OnData { get; set; }

    /// <summary>
    /// 获取或设置在双击工作表中的任何位置时运行的宏名称
    /// </summary>
    string OnDoubleClick { get; set; }


    /// <summary>
    /// 获取或设置一个值，该值指示是否显示自动分页符
    /// </summary>
    bool DisplayAutomaticPageBreaks { get; set; }

    /// <summary>
    /// 获取工作表的代码名称
    /// 对应 Worksheet.CodeName 属性
    /// </summary>
    string? CodeName { get; }

    /// <summary>
    /// 获取工作表的下一个工作表
    /// 对应 Worksheet.Next 属性
    /// </summary>
    IExcelWorksheet? Next { get; }

    /// <summary>
    /// 获取工作表的上一个工作表
    /// 对应 Worksheet.Previous 属性
    /// </summary>
    IExcelWorksheet? Previous { get; }

    #endregion

    #region 区域访问

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// 对应 Worksheet.Range 属性
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    IExcelRange? Range(object? cell1, object? cell2 = null);

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    IExcelRange? this[int row, int column] { get; }

    /// <summary>
    /// 获取工作表中指定地址的单元格对象
    /// </summary>
    IExcelRange? this[string address] { get; }

    /// <summary>
    /// 获取工作表中指定地址的单元格对象
    /// </summary>
    /// <param name="begin">开始地址：如A1</param>
    /// <param name="end">结束地址：如A5</param>
    IExcelRange? this[string begin, string end] { get; }

    /// <summary>
    /// 获取工作表的所有行
    /// </summary>
    IExcelRange? Rows { get; }

    /// <summary>
    /// 获取工作表的所有列
    /// </summary>
    IExcelRange? Columns { get; }

    /// <summary>
    /// 获取工作表中指定行的区域对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>行区域对象</returns>
    IExcelRange? GetRow(int row);

    /// <summary>
    /// 获取工作表中指定列的区域对象
    /// </summary>
    /// <param name="column">列号</param>
    /// <returns>列区域对象</returns>
    IExcelRange? GetColumn(int column);

    /// <summary>
    /// 获取工作表的已使用区域
    /// 对应 Worksheet.UsedRange 属性
    /// </summary>
    IExcelRange? UsedRange { get; }

    /// <summary>
    /// 获取工作表的整个区域
    /// </summary>
    IExcelRange? AllRange { get; }

    #endregion

    #region 形状和图表

    /// <summary>
    /// 获取工作表的列表对象集合
    /// 对应 Worksheet.ListObjects 属性
    /// </summary>
    IExcelListObjects? ListObjects { get; }

    /// <summary>
    /// 获取工作表的图片集合
    /// </summary>
    IExcelPictures? Pictures { get; }

    /// <summary>
    /// 获取工作表的评论集合
    /// </summary>
    IExcelComments? Comments { get; }

    #endregion

    #region 数据操作

    IExcelQueryTables? QueryTables { get; }

    /// <summary>
    /// 获取或设置工作表的默认行高
    /// </summary>
    double DefaultRowHeight { get; set; }

    /// <summary>
    /// 获取或设置工作表的默认列宽
    /// </summary>
    double DefaultColumnWidth { get; set; }

    /// <summary>
    /// 获取或设置是否启用自动筛选
    /// </summary>
    bool EnableAutoFilter { get; set; }

    /// <summary>
    /// 获取或设置是否启用计算器
    /// </summary>
    bool EnableCalculation { get; set; }

    /// <summary>
    /// 获取或设置是否显示页面布局
    /// </summary>
    bool DisplayPageLayout { get; set; }
    #endregion

    #region 操作方法  

    /// <summary>
    /// 粘贴剪贴板内容到指定单元格
    /// </summary>
    /// <param name="destinationCell">目标单元格</param>
    /// <param name="link"></param>
    void Paste(IExcelRange destinationCell, bool? link = null);

    /// <summary>
    /// 粘贴剪贴板内容到指定区域
    /// </summary>
    /// <param name="startRow">起始行</param>
    /// <param name="startColumn">起始列</param>
    void PasteToPosition(int startRow, int startColumn);

    /// <summary>
    /// 通用特殊粘贴方法
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    /// <param name="pasteType">粘贴类型</param>
    /// <param name="skipBlanks">是否跳过空白单元格</param>
    /// <param name="transpose">是否转置</param>
    void PasteSpecial(IExcelRange destinationRange,
                          XlPasteType pasteType = XlPasteType.xlPasteAll,
                          bool skipBlanks = false,
                          bool transpose = false);

    /// <summary>
    /// 特殊粘贴 - 只粘贴值
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    void PasteValues(IExcelRange destinationRange);

    /// <summary>
    /// 特殊粘贴 - 只粘贴格式
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    void PasteFormats(IExcelRange destinationRange);

    /// <summary>
    /// 特殊粘贴 - 粘贴公式
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    void PasteFormulas(IExcelRange destinationRange);

    /// <summary>
    /// 特殊粘贴 - 粘贴列宽
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    void PasteColumnWidths(IExcelRange destinationRange);

    /// <summary>
    /// 特殊粘贴 - 执行计算操作
    /// </summary>
    /// <param name="destinationRange">目标区域</param>
    /// <param name="operation">计算操作类型</param>
    void PasteWithOperation(IExcelRange destinationRange, XlPasteSpecialOperation operation);


    /// <summary>
    /// 获取工作表中指定索引的数据透视表对象
    /// </summary>
    /// <param name="index">数据透视表索引</param>
    /// <returns>数据透视表对象</returns>
    IExcelPivotTable? PivotTables(int index);

    /// <summary>
    /// 获取工作表中所有数据透视表的集合
    /// </summary>
    /// <returns>数据透视表集合对象</returns>
    IExcelPivotTables? PivotTables();

    /// <summary>
    /// 获取工作表所在工作簿的数据透视表缓存集合
    /// </summary>
    /// <returns>数据透视表缓存集合对象，如果获取失败则返回null</returns>
    IExcelPivotCaches? PivotCaches();

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    IExcelChartObjects? ChartObjects();

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    IExcelChartObject? ChartObjects(int index);

    /// <summary>
    /// 获取工作表的图表对象集合
    /// </summary>
    IExcelChartObject? ChartObjects(string name);

    /// <summary>
    /// 将工作表导出为固定格式文件（如PDF或XPS）
    /// 对应 Worksheet.ExportAsFixedFormat 方法
    /// </summary>
    /// <param name="type">导出文件格式类型（PDF或XPS）</param>
    /// <param name="filename">导出文件的完整路径和文件名</param>
    /// <param name="quality">导出质量，可影响文件大小和质量</param>
    /// <param name="includeDocProperties">是否包含文档属性信息</param>
    /// <param name="ignorePrintAreas">是否忽略打印区域设置，如果为true则导出整个工作表</param>
    /// <param name="from">起始页码，指定要导出的起始页面</param>
    /// <param name="to">结束页码，指定要导出的结束页面</param>
    /// <param name="openAfterPublish">导出完成后是否自动打开文件</param>
    void ExportAsFixedFormat(
        XlFixedFormatType type,
        string filename,
        XlFixedFormatQuality? quality = null,
        bool? includeDocProperties = null,
        bool? ignorePrintAreas = null,
        int? from = null,
        int? to = null,
        bool? openAfterPublish = null);

    /// <summary>
    /// 重命名工作表
    /// </summary>
    /// <param name="newName">新名称</param>
    void Rename(string newName);


    #endregion

    #region 高级功能

    /// <summary>
    /// 在活动工作表中创建指定范围
    /// </summary>
    IExcelRange CreateRange(string address);

    /// <summary>
    /// 在活动工作表中获取指定行
    /// </summary>
    IExcelRows GetRows(int startRow, int endRow = -1);

    /// <summary>
    /// 在活动工作表中获取指定列
    /// </summary>
    IExcelColumns GetColumns(string startColumn, string endColumn = "");


    /// <summary>
    /// 重置工作表中的所有分页符
    /// 清除所有手动添加的水平和垂直分页符，恢复到默认的分页符设置
    /// </summary>
    void ResetAllPageBreaks();

    /// <summary>
    /// 计算工作表中的所有公式
    /// 对应 Worksheet.Calculate 方法
    /// </summary>
    void Calculate();

    /// <summary>
    /// 重新计算工作表
    /// </summary>
    void Recalculate();

    /// <summary>
    /// 清除工作表格式
    /// </summary>
    void ClearFormats();

    /// <summary>
    /// 清除工作表注释
    /// </summary>
    void ClearComments();

    /// <summary>
    /// 清除工作表超链接
    /// </summary>
    void ClearHyperlinks();

    /// <summary>
    /// 自动调整列宽
    /// </summary>
    void AutoFitColumns();

    /// <summary>
    /// 自动调整行高
    /// </summary>
    void AutoFitRows();

    #endregion

    /// <summary>
    /// 保护工作表
    /// 对应 Worksheet.Protect 方法
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="drawingObjects">是否保护图形对象</param>
    /// <param name="contents">是否保护内容</param>
    /// <param name="scenarios">是否保护方案</param>
    /// <param name="userInterfaceOnly">是否仅保护用户界面</param>
    /// <param name="allowFormattingCells">是否允许格式化单元格</param>
    /// <param name="allowFormattingColumns">是否允许格式化列</param>
    /// <param name="allowFormattingRows">是否允许格式化行</param>
    /// <param name="allowInsertingColumns">是否允许插入列</param>
    /// <param name="allowInsertingRows">是否允许插入行</param>
    /// <param name="allowInsertingHyperlinks">是否允许插入超链接</param>
    /// <param name="allowDeletingColumns">是否允许删除列</param>
    /// <param name="allowDeletingRows">是否允许删除行</param>
    /// <param name="allowSorting">是否允许排序</param>
    /// <param name="allowFiltering">是否允许筛选</param>
    /// <param name="allowUsingPivotTables">是否允许使用透视表</param>
    void Protect(string password = "", bool drawingObjects = true, bool contents = true,
                bool scenarios = true, bool userInterfaceOnly = false,
                bool allowFormattingCells = true, bool allowFormattingColumns = true,
                bool allowFormattingRows = true, bool allowInsertingColumns = true,
                bool allowInsertingRows = true, bool allowInsertingHyperlinks = true,
                bool allowDeletingColumns = true, bool allowDeletingRows = true,
                bool allowSorting = true, bool allowFiltering = true,
                bool allowUsingPivotTables = true);

    #region 事件

    /// <summary>
    /// 当工作表内容发生改变时触发
    /// </summary>
    event ChangeEventHandler Change;

    /// <summary>
    /// 当工作表选择区域发生改变时触发
    /// </summary>
    event SelectionChangeEventHandler SelectionChange;

    /// <summary>
    /// 当工作表被激活时触发
    /// </summary>
    event ActivateEventHandler SheetActivate;

    /// <summary>
    /// 当工作表被取消激活时触发
    /// </summary>
    event DeactivateEventHandler SheetDeactivate;

    /// <summary>
    /// 当工作表被双击时触发
    /// </summary>
    event BeforeDoubleClickEventHandler BeforeDoubleClick;

    /// <summary>
    /// 当工作表被右键单击时触发
    /// </summary>
    event BeforeRightClickEventHandler BeforeRightClick;

    /// <summary>
    /// 当工作表计算完成后触发
    /// </summary>
    event CalculateEventHandler SheetCalculate;

    /// <summary>
    /// 在工作表被删除之前触发
    /// </summary>
    event BeforeDeleteEventHandler BeforeDelete;

    /// <summary>
    /// 当数据透视表发生更改时同步触发
    /// </summary>
    event PivotTableChangeSyncEventHandler PivotTableChangeSync;
    #endregion
}