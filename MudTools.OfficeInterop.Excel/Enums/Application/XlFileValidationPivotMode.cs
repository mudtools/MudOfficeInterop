
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定如何验证数据透视表的数据缓存。
/// </summary>
public enum XlFileValidationPivotMode
{
    /// <summary>
    /// 按照 PivotOptions 注册表设置的指示来验证数据缓存（默认）。
    /// </summary>
    xlFileValidationPivotDefault,

    /// <summary>
    /// 忽略注册表设置，验证所有数据缓存。
    /// </summary>
    xlFileValidationPivotRun,

    /// <summary>
    /// 不验证数据缓存的内容。
    /// </summary>
    xlFileValidationPivotSkip
}