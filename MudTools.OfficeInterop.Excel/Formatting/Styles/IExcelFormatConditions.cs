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
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelFormatConditions : IOfficeObject<IExcelFormatConditions>, IEnumerable<IExcelFormatCondition>, IDisposable
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
    [ComPropertyWrap(NeedConvert = true)]
    IExcelFormatCondition? this[int index] { get; }

    /// <summary>
    /// 根据名称获取指定的条件格式规则对象
    /// </summary>
    /// <param name="name">条件格式规则的名称</param>
    /// <returns>与指定名称对应的条件格式规则对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelFormatCondition? this[string name] { get; }

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
    /// <param name="strOperator">字符串</param>
    /// <param name="textOperator">文本</param>
    /// <param name="dateOperator">日期</param>
    /// <param name="scopeType">范围类型</param>
    /// <returns>新创建的条件格式规则对象</returns>
    [IgnoreGenerator]
    IExcelFormatCondition? Add(
         XlFormatConditionType type,
         XlFormatConditionOperator? @operator,
         object? formula1 = null,
         object? formula2 = null,
         object? strOperator = null,
         object? textOperator = null,
         object? dateOperator = null,
         object? scopeType = null);


    /// <summary>
    /// 向集合中添加新的条件格式规则 (xlExpression)
    /// 对应 FormatConditions.Add 方法
    /// </summary>
    /// <param name="formula">条件公式</param>
    /// <returns>新创建的条件格式规则对象</returns>
    [IgnoreGenerator]
    IExcelFormatCondition? AddExpression(string formula);

    /// <summary>
    /// 向集合中添加新的数据条条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlDatabar)
    /// </summary>
    /// <returns>新创建的条件格式规则对象</returns>
    [ValueConvert]
    IExcelDatabar? AddDatabar();

    /// <summary>
    /// 向集合中添加新的图标集条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlIconSet)
    /// </summary>
    /// <returns>新创建的条件格式规则对象</returns>
    [ValueConvert]
    IExcelIconSetCondition? AddIconSetCondition();

    /// <summary>
    /// 向集合中添加新的唯一值/重复值条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlUniqueValues)
    /// </summary>
    /// <returns>新创建的条件格式规则对象</returns>
    [ValueConvert]
    IExcelUniqueValues? AddUniqueValues();

    /// <summary>
    /// 向集合中添加新的TOP N条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlTop10)
    /// </summary>
    /// <returns>新创建的条件格式规则对象</returns>
    [ValueConvert]
    IExcelTop10? AddTop10();

    /// <summary>
    /// 向集合中添加新的高于平均值条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlAboveAverage)
    /// </summary>
    /// <returns>新创建的高于平均值条件格式规则对象</returns>
    [ValueConvert]
    IExcelAboveAverage AddAboveAverage();

    /// <summary>
    /// 向集合中添加新的色阶条件格式规则
    /// 对应 FormatConditions.Add 方法 (使用 XlFormatConditionType.xlColorScale)
    /// </summary>
    /// <param name="ColorScaleType">色阶类型</param>
    /// <returns>新创建的色阶条件格式规则对象</returns>
    [ValueConvert]
    IExcelColorScale AddColorScale(int ColorScaleType);

    /// <summary>
    /// 删除条件格式规则
    /// </summary>
    void Delete();
    #endregion
}
