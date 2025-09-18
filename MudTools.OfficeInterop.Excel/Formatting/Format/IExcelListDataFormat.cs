//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 表格列（ListColumn）的数据格式信息的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ListDataFormat
/// 提供数据类型、校验、默认值等只读属性。
/// </summary>
public interface IExcelListDataFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 ListColumn）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取该列数据的默认值（如果未设置则返回 null）。
    /// </summary>
    object DefaultValue { get; }

    /// <summary>
    /// 获取该列是否允许空值。
    /// true = 允许空值，false = 不允许。
    /// </summary>
    bool AllowFillIn { get; }

    /// <summary>
    /// 获取该列是否为“必需”字段（即不允许空值）。
    /// </summary>
    bool Required { get; }

    /// <summary>
    /// 获取该列数据类型（如文本、数字、日期等）。
    /// 使用 Excel.XlListDataType 枚举。
    /// </summary>
    XlListDataType Type { get; }

    int MaxCharacters { get; }

    /// <summary>
    /// 获取该列数据校验的最小值（仅对数字/日期类型有效）。
    /// 如果未设置或类型不支持，返回 null。
    /// </summary>
    object MinNumber { get; }

    /// <summary>
    /// 获取该列数据校验的最大值（仅对数字/日期类型有效）。
    /// 如果未设置或类型不支持，返回 null。
    /// </summary>
    object MaxNumber { get; }

    /// <summary>
    /// 获取该列是否启用“只读”模式（用户不能编辑）。
    /// </summary>
    bool ReadOnly { get; }


    int DecimalPlaces { get; }

}