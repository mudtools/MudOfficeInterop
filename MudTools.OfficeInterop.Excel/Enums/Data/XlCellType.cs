//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 单元格类型枚举
/// 用于指定特定类型的单元格，通常在查找或选择操作中使用
/// </summary>
public enum XlCellType
{
    /// <summary>
    /// 空单元格
    /// 不包含任何数据或格式的单元格
    /// </summary>
    xlCellTypeBlanks = 4,

    /// <summary>
    /// 包含常量的单元格
    /// 包含文本、数字、日期、时间或逻辑值等常量数据的单元格
    /// </summary>
    xlCellTypeConstants = 2,

    /// <summary>
    /// 包含公式的单元格
    /// 包含公式而非常量值的单元格
    /// </summary>
    xlCellTypeFormulas = -4123,

    /// <summary>
    /// 最后一个单元格
    /// 工作表中包含数据或格式的最右下角的单元格
    /// </summary>
    xlCellTypeLastCell = 11,

    /// <summary>
    /// 包含批注的单元格
    /// 包含单元格批注的单元格
    /// </summary>
    xlCellTypeComments = -4144,

    /// <summary>
    /// 可见单元格
    /// 不被隐藏的单元格（在筛选操作后可见的单元格）
    /// </summary>
    xlCellTypeVisible = 12,

    /// <summary>
    /// 所有带格式条件的单元格
    /// 应用了条件格式的所有单元格
    /// </summary>
    xlCellTypeAllFormatConditions = -4172,

    /// <summary>
    /// 相同格式条件的单元格
    /// 具有相同条件格式的单元格
    /// </summary>
    xlCellTypeSameFormatConditions = -4173,

    /// <summary>
    /// 所有带数据验证的单元格
    /// 应用了数据验证规则的所有单元格
    /// </summary>
    xlCellTypeAllValidation = -4174,

    /// <summary>
    /// 相同数据验证的单元格
    /// 具有相同数据验证规则的单元格
    /// </summary>
    xlCellTypeSameValidation = -4175
}