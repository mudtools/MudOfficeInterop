//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 粘贴类型枚举
/// 用于指定执行粘贴操作时要粘贴的内容类型
/// </summary>
public enum PasteType
{
    /// <summary>
    /// 粘贴所有内容，包括数值、公式、格式、批注等
    /// </summary>
    All = -4104,
    
    /// <summary>
    /// 粘贴除边框外的所有内容
    /// </summary>
    AllExceptBorders = 7,
    
    /// <summary>
    /// 粘贴所有内容并合并条件格式
    /// </summary>
    AllMergingConditionalFormats = 14,
    
    /// <summary>
    /// 使用源主题粘贴所有内容
    /// </summary>
    AllUsingSourceTheme = 13,
    
    /// <summary>
    /// 仅粘贴列宽
    /// </summary>
    ColumnWidths = 8,
    
    /// <summary>
    /// 仅粘贴批注
    /// </summary>
    Comments = -4144,
    
    /// <summary>
    /// 仅粘贴格式
    /// </summary>
    Formats = -4122,
    
    /// <summary>
    /// 仅粘贴公式
    /// </summary>
    Formulas = -4123,
    
    /// <summary>
    /// 粘贴公式和数字格式
    /// </summary>
    FormulasAndNumberFormats = 11,
    
    /// <summary>
    /// 仅粘贴数据验证规则
    /// </summary>
    Validation = 6,
    
    /// <summary>
    /// 仅粘贴数值
    /// </summary>
    Values = -4163,
    
    /// <summary>
    /// 粘贴数值和数字格式
    /// </summary>
    ValuesAndNumberFormats = 12
}
