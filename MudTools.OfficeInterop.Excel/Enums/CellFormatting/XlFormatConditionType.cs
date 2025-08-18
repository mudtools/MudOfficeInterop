//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定条件格式的类型
/// </summary>
public enum XlFormatConditionType
{
    /// <summary>
    /// 基于单元格值的条件格式
    /// </summary>
    xlCellValue = 1,
    
    /// <summary>
    /// 基于公式表达式的条件格式
    /// </summary>
    xlExpression = 2,
    
    /// <summary>
    /// 颜色刻度条件格式
    /// </summary>
    xlColorScale = 3,
    
    /// <summary>
    /// 数据条条件格式
    /// </summary>
    xlDatabar = 4,
    
    /// <summary>
    /// 前10项条件格式
    /// </summary>
    xlTop10 = 5,
    
    /// <summary>
    /// 图标集条件格式
    /// </summary>
    xlIconSets = 6,
    
    /// <summary>
    /// 唯一值条件格式
    /// </summary>
    xlUniqueValues = 8,
    
    /// <summary>
    /// 文本字符串条件格式
    /// </summary>
    xlTextString = 9,
    
    /// <summary>
    /// 空白单元格条件格式
    /// </summary>
    xlBlanksCondition = 10,
    
    /// <summary>
    /// 时间段条件格式
    /// </summary>
    xlTimePeriod = 11,
    
    /// <summary>
    /// 高于平均值条件格式
    /// </summary>
    xlAboveAverageCondition = 12,
    
    /// <summary>
    /// 非空白单元格条件格式
    /// </summary>
    xlNoBlanksCondition = 13,
    
    /// <summary>
    /// 错误单元格条件格式
    /// </summary>
    xlErrorsCondition = 16,
    
    /// <summary>
    /// 非错误单元格条件格式
    /// </summary>
    xlNoErrorsCondition = 17
}