//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Names 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Names 的安全访问和操作
/// </summary>
public interface IExcelNames : IEnumerable<IExcelName>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取名称集合中的名称数量
    /// 对应 Names.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的名称对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">名称索引（从1开始）</param>
    /// <returns>名称对象</returns>
    IExcelName this[int index] { get; }

    /// <summary>
    /// 获取指定名称的名称对象
    /// </summary>
    /// <param name="name">名称</param>
    /// <returns>名称对象</returns>
    IExcelName this[string name] { get; }

    /// <summary>
    /// 获取名称集合所在的父对象（通常是工作簿或工作表）
    /// 对应 Names.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取名称集合所在的Application对象
    /// 对应 Names.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 添加新的名称
    /// 对应 Names.Add 方法
    /// </summary>
    /// <param name="name">名称</param>
    /// <param name="refersTo">引用</param>
    /// <param name="visible">是否可见</param>
    /// <param name="macroType">宏类型</param>
    /// <param name="shortcutKey">快捷键</param>
    /// <param name="category">类别</param>
    /// <param name="nameLocal">本地名称</param>
    /// <param name="refersToLocal">本地引用</param>
    /// <param name="categoryLocal">本地类别</param>
    /// <param name="refersToR1C1">R1C1引用</param>
    /// <param name="refersToR1C1Local">本地R1C1引用</param>
    /// <returns>新创建的名称对象</returns>
    IExcelName? Add(string name, object? refersTo = null, bool visible = true,
                         int macroType = 0, string shortcutKey = "", object? category = null,
                         string nameLocal = "", object? refersToLocal = null, object? categoryLocal = null,
                         string refersToR1C1 = "", string refersToR1C1Local = "");


    /// <summary>
    /// 基于区域创建名称
    /// </summary>
    /// <param name="range">区域对象</param>
    /// <param name="name">名称</param>
    /// <param name="useColumnNames">是否使用列名</param>
    /// <param name="useRowNames">是否使用行名</param>
    /// <returns>创建的名称对象</returns>
    IExcelName? CreateFromRange(IExcelRange range, string name = "",
                              bool useColumnNames = false, bool useRowNames = false);

    /// <summary>
    /// 创建工作表名称
    /// </summary>
    /// <param name="worksheet">工作表对象</param>
    /// <param name="name">名称</param>
    /// <returns>创建的名称对象</returns>
    IExcelName? CreateWorksheetName(IExcelWorksheet worksheet, string name = "");

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找
    /// </summary>
    /// <param name="name">名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的名称数组</returns>
    IExcelName[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据引用查找
    /// </summary>
    /// <param name="refersTo">引用</param>
    /// <returns>匹配的名称数组</returns>
    IExcelName[] FindByRefersTo(string refersTo);

    /// <summary>
    /// 根据可见性查找
    /// </summary>
    /// <param name="visible">可见性</param>
    /// <returns>匹配的名称数组</returns>
    IExcelName[] FindByVisibility(bool visible);

    /// <summary>
    /// 根据类别查找
    /// </summary>
    /// <param name="category">类别</param>
    /// <returns>匹配的名称数组</returns>
    IExcelName[] FindByCategory(string category);

    /// <summary>
    /// 获取可见的名称
    /// </summary>
    /// <returns>可见名称数组</returns>
    IExcelName[] GetVisibleNames();

    /// <summary>
    /// 获取隐藏的名称
    /// </summary>
    /// <returns>隐藏名称数组</returns>
    IExcelName[] GetHiddenNames();

    /// <summary>
    /// 获取工作簿级别的名称
    /// </summary>
    /// <returns>工作簿级别名称数组</returns>
    IExcelName[] GetWorkbookNames();

    /// <summary>
    /// 获取工作表级别的名称
    /// </summary>
    /// <returns>工作表级别名称数组</returns>
    IExcelName[] GetWorksheetNames();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有名称
    /// 对应 Names.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的名称
    /// </summary>
    /// <param name="index">要删除的名称索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的名称
    /// </summary>
    /// <param name="name">要删除的名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的名称对象
    /// </summary>
    /// <param name="nameObject">要删除的名称对象</param>
    void Delete(IExcelName nameObject);

    /// <summary>
    /// 批量删除名称
    /// </summary>
    /// <param name="names">要删除的名称数组</param>
    void DeleteRange(string[] names);

    /// <summary>
    /// 选择所有名称
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 取消选择所有名称
    /// </summary>
    void DeselectAll();

    /// <summary>
    /// 刷新所有名称
    /// </summary>
    void Refresh();

    #endregion

    #region 导出和导入

    /// <summary>
    /// 导出所有名称到文本文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="includeHidden">是否包含隐藏名称</param>
    /// <returns>是否导出成功</returns>
    bool ExportToText(string filename, bool includeHidden = false);

    /// <summary>
    /// 从文本文件导入名称
    /// </summary>
    /// <param name="filename">导入文件路径</param>
    /// <returns>成功导入的名称数量</returns>
    int ImportFromText(string filename);

    #endregion


    #region 高级功能

    /// <summary>
    /// 获取活动名称
    /// </summary>
    /// <returns>活动名称对象</returns>
    IExcelName ActiveName { get; }
    #endregion
}