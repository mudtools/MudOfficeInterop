//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel FormatConditions 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.FormatConditions 的安全访问和操作
/// </summary>
public interface IExcelFormatConditions : IEnumerable<IExcelFormatCondition>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取条件格式规则集合中的规则数量
    /// 对应 FormatConditions.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的条件格式规则对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">规则索引（从1开始）</param>
    /// <returns>条件格式规则对象</returns>
    IExcelFormatCondition? this[int index] { get; }

    /// <summary>
    /// 获取条件格式集合所在的父对象（通常是 Range）
    /// 对应 FormatConditions.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件格式集合所在的Application对象
    /// 对应 FormatConditions.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的条件格式规则 (xlCellValue)
    /// 对应 FormatConditions.Add 方法
    /// </summary>
    /// <param name="type">条件类型</param>
    /// <param name="operator">比较操作符</param>
    /// <param name="formula1">公式1</param>
    /// <param name="formula2">公式2</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? Add(
         XlFormatConditionType type,
         XlFormatConditionOperator? @operator,
         object? formula1 = null,
         object? formula2 = null,
         object? @string = null,
         object? textOperator = null,
         object? dateOperator = null,
         object? scopeType = null);

    /// <summary>
    /// 向集合中添加新的条件格式规则 (xlExpression)
    /// 对应 FormatConditions.Add 方法
    /// </summary>
    /// <param name="formula">条件公式</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? AddExpression(string formula);

    /// <summary>
    /// 向集合中添加新的颜色刻度条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlColorScale)
    /// </summary>
    /// <param name="colorScaleType">颜色刻度类型 (例如 3 for ThreeColorScale)</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? AddColorScale(int colorScaleType);

    /// <summary>
    /// 向集合中添加新的数据条条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlDatabar)
    /// </summary>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelDataBar? AddDatabar();

    /// <summary>
    /// 向集合中添加新的图标集条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlIconSet)
    /// </summary>
    /// <param name="iconSet">图标集类型</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? AddIconSetCondition(XlIconSet iconSet);

    /// <summary>
    /// 向集合中添加新的唯一值/重复值条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlUniqueValues)
    /// </summary>
    /// <param name="showUnique">true为唯一值，false为重复值</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? AddUniqueValues(bool showUnique);

    /// <summary>
    /// 向集合中添加新的TOP N条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlTop10)
    /// </summary>
    /// <param name="rank">排名 (1-1000)</param>
    /// <param name="aboveAverage">true为高于平均值，false为低于平均值</param>
    /// <param name="percent">是否按百分比计算</param>
    /// <returns>新创建的条件格式规则对象</returns>
    IExcelFormatCondition? AddTop10(int rank, bool aboveAverage = true, bool percent = false);
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据条件类型查找规则
    /// </summary>
    /// <param name="type">条件类型</param>
    /// <returns>匹配的规则数组</returns>
    IExcelFormatCondition[] FindByType(XlFormatConditionType type);
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除条件格式规则
    /// </summary>
    void Delete();

    /// <summary>
    /// 删除指定索引的条件格式规则
    /// </summary>
    /// <param name="index">要删除的规则索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的条件格式规则对象
    /// </summary>
    /// <param name="condition">要删除的条件格式规则对象</param>
    void Delete(IExcelFormatCondition condition);

    /// <summary>
    /// 批量删除条件格式规则
    /// </summary>
    /// <param name="indices">要删除的规则索引数组</param>
    void DeleteRange(int[] indices);

    #endregion
}
