//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel工作表集合的公共接口
/// </summary>
public interface IExcelComSheets : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取工作表集合中的工作表数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取工作表集合所在的父对象（通常是 Workbook）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取工作表集合所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取工作表的水平分页符集合
    /// </summary>
    IExcelHPageBreaks? HPageBreaks { get; }

    /// <summary>
    /// 获取工作表的垂直分页符集合
    /// </summary>
    IExcelVPageBreaks? VPageBreaks { get; }

    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的工作表
    /// 对应 Sheets.Add 方法
    /// </summary>
    /// <param name="before">在哪个工作表之前插入</param>
    /// <param name="after">在哪个工作表之后插入</param>
    /// <param name="count">要添加的工作表数量</param>
    /// <returns>新创建的工作表对象 (或第一个，如果添加了多个)</returns>
    IExcelWorksheet? AddSheet(
        IExcelComSheet? before = null,
        IExcelComSheet? after = null,
        int? count = 1);

    /// <summary>
    /// 向集合中添加新的工作表
    /// 对应 Sheets.Add 方法
    /// </summary>
    /// <param name="before">在哪个工作表之前插入</param>
    /// <param name="after">在哪个工作表之后插入</param>
    /// <param name="count">要添加的工作表数量</param>
    /// <param name="type">工作表类型</param>
    /// <returns>新创建的工作表对象 (或第一个，如果添加了多个)</returns>
    IExcelComSheet? Add(
        IExcelComSheet? before = null,
        IExcelComSheet? after = null,
        int? count = 1,
        XlSheetType? type = null);

    /// <summary>
    /// 从文件复制工作表到此集合
    /// </summary>
    /// <param name="filename">源文件路径</param>
    /// <param name="sheetName">源工作表名称</param>
    /// <param name="before">在哪个工作表之前插入</param>
    /// <param name="after">在哪个工作表之后插入</param>
    /// <returns>新创建的工作表对象</returns>
    IExcelComSheet? CreateFromTemplate(
        string filename,
        string sheetName,
        IExcelComSheet? before = null,
        IExcelComSheet? after = null);
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 获取可见的工作表
    /// </summary>
    /// <returns>可见工作表数组</returns>
    IExcelComSheet[] GetVisibleSheets();

    /// <summary>
    /// 获取隐藏的工作表
    /// </summary>
    /// <returns>隐藏工作表数组</returns>
    IExcelComSheet[] GetHiddenSheets();

    /// <summary>
    /// 获取非常隐藏的工作表 (xlSheetVeryHidden)
    /// </summary>
    /// <returns>非常隐藏工作表数组</returns>
    IExcelComSheet[] GetVeryHiddenSheets();

    /// <summary>
    /// 获取受保护的工作表
    /// </summary>
    /// <returns>受保护工作表数组</returns>
    IExcelComSheet[] GetProtectedSheets();

    /// <summary>
    /// 根据名称查找工作表
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的工作表数组</returns>
    IExcelComSheet[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据类型查找工作表 (例如: xlWorksheet, xlChart, xlExcel4MacroSheet, xlExcel4IntlMacroSheet)
    /// </summary>
    /// <param name="type">工作表类型</param>
    /// <returns>匹配的工作表数组</returns>
    IExcelComSheet[] FindByType(XlSheetType type);

    /// <summary>
    /// 根据索引范围查找工作表
    /// </summary>
    /// <param name="startIndex">起始索引</param>
    /// <param name="endIndex">结束索引</param>
    /// <returns>匹配的工作表数组</returns>
    IExcelComSheet[] FindByIndexRange(int startIndex, int endIndex);

    #endregion

    #region 操作方法

    /// <summary>
    /// 添加新工作表
    /// </summary>
    /// <param name="options">添加选项</param>
    /// <returns>新创建的工作表</returns>
    IExcelComSheet AddSheet(AddSheetOptions options);

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="source">源工作表</param>
    /// <param name="options">复制选项</param>
    /// <returns>新创建的工作表副本</returns>
    IExcelComSheet CopySheet(IExcelComSheet source, CopySheetOptions options);

    /// <summary>
    /// 删除指定索引的工作表
    /// </summary>
    /// <param name="index">要删除的工作表索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的工作表
    /// </summary>
    /// <param name="name">要删除的工作表名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的工作表对象
    /// </summary>
    /// <param name="sheet">要删除的工作表对象</param>
    void Delete(IExcelComSheet sheet);

    /// <summary>
    /// 批量删除工作表
    /// </summary>
    /// <param name="indices">要删除的工作表索引数组</param>
    void DeleteRange(int[] indices);

    /// <summary>
    /// 批量删除工作表
    /// </summary>
    /// <param name="names">要删除的工作表名称数组</param>
    void DeleteRange(string[] names);

    /// <summary>
    /// 选择多个工作表
    /// </summary>
    /// <param name="worksheetNames">工作表名称数组</param>
    void Select(params string[] worksheetNames);

    #endregion

    #region 高级功能

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    IExcelComSheet? ActiveWorksheet { get; }

    /// <summary>
    /// 打印所有工作表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOutAll(bool preview = false);

    /// <summary>
    /// 计算所有工作表
    /// </summary>
    void Calculate();

    /// <summary>
    /// 刷新所有工作表
    /// </summary>
    void RefreshAll();

    /// <summary>
    /// 保护所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    void ProtectAll(string password = "");

    /// <summary>
    /// 取消保护所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    void UnprotectAll(string password = "");

    #endregion
}

public class CopySheetOptions
{
    /// <summary>复制到该工作表之前</summary>
    public IExcelWorksheet Before { get; set; }

    /// <summary>复制到该工作表之后</summary>
    public IExcelWorksheet After { get; set; }

    /// <summary>目标工作簿（跨工作簿复制）</summary>
    public IExcelWorkbook TargetWorkbook { get; set; }

    /// <summary>是否仅复制值（不包含格式和公式）</summary>
    public bool ValuesOnly { get; set; }
}

/// <summary>
/// 添加工作表选项
/// </summary>
public class AddSheetOptions
{
    /// <summary>添加到该工作表之前</summary>
    public IExcelWorksheet Before { get; set; }

    /// <summary>添加到该工作表之后</summary>
    public IExcelWorksheet After { get; set; }

    /// <summary>添加数量（默认1）</summary>
    public int Count { get; set; } = 1;

    /// <summary>工作表类型（默认普通工作表）</summary>
    public XlSheetType Type { get; set; } = XlSheetType.xlWorksheet;

    /// <summary>基于现有工作表模板创建</summary>
    public IExcelWorksheet Template { get; set; }

    /// <summary>新工作表名称</summary>
    public string Name { get; set; }
}