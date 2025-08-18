//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 合并计算函数枚举
/// 用于指定数据透视表或合并计算中使用的汇总函数
/// </summary>
public enum XlConsolidationFunction
{
    /// <summary>
    /// 平均值
    /// 计算数据的算术平均值
    /// </summary>
    xlAverage = -4106,
    
    /// <summary>
    /// 计数
    /// 计算包含数字和文本值的单元格数量
    /// </summary>
    xlCount = -4112,
    
    /// <summary>
    /// 数字计数
    /// 计算包含数字值的单元格数量
    /// </summary>
    xlCountNums = -4113,
    
    /// <summary>
    /// 最大值
    /// 返回数据中的最大值
    /// </summary>
    xlMax = -4136,
    
    /// <summary>
    /// 最小值
    /// 返回数据中的最小值
    /// </summary>
    xlMin = -4139,
    
    /// <summary>
    /// 乘积
    /// 计算所有数值的乘积
    /// </summary>
    xlProduct = -4149,
    
    /// <summary>
    /// 标准偏差
    /// 基于样本估算标准偏差
    /// </summary>
    xlStDev = -4155,
    
    /// <summary>
    /// 总体标准偏差
    /// 计算基于整个样本总体的标准偏差
    /// </summary>
    xlStDevP = -4156,
    
    /// <summary>
    /// 求和
    /// 计算所有数值的总和
    /// </summary>
    xlSum = -4157,
    
    /// <summary>
    /// 方差
    /// 基于样本估算方差
    /// </summary>
    xlVar = -4164,
    
    /// <summary>
    /// 总体方差
    /// 计算基于整个样本总体的方差
    /// </summary>
    xlVarP = -4165,
    
    /// <summary>
    /// 未知函数
    /// 表示未知或不支持的函数类型
    /// </summary>
    xlUnknown = 1000,
    
    /// <summary>
    /// 不重复计数
    /// 计算不重复值的数量
    /// </summary>
    xlDistinctCount = 11
}
