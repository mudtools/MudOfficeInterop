//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示在参数查询中使用的单个参数。Parameter对象是Parameters集合的成员。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelParameter : IOfficeObject<IExcelParameter, MsExcel.Parameter>, IDisposable
{
    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示Excel应用程序的Application对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置指定查询参数的数据类型。
    /// </summary>
    XlParameterDataType DataType { get; set; }

    /// <summary>
    /// 获取参数类型。
    /// </summary>
    XlParameterType Type { get; }

    /// <summary>
    /// 获取在参数查询中提示用户输入参数值的短语。
    /// </summary>
    string PromptString { get; }

    /// <summary>
    /// 获取参数值。
    /// </summary>
    object Value { get; }

    /// <summary>
    /// 获取表示包含指定查询参数值的单元格的Range对象。
    /// </summary>
    IExcelRange? SourceRange { get; }

    /// <summary>
    /// 获取或设置参数的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 为指定的查询表定义参数。
    /// </summary>
    /// <param name="type">必需。参数类型。</param>
    /// <param name="value">必需。指定参数的值，如Type参数描述所示。</param>
    void SetParam(XlParameterType type, object value);

    /// <summary>
    /// 获取或设置一个布尔值，表示每当更改参数查询的参数值时，是否刷新指定的查询表。
    /// </summary>
    bool RefreshOnChange { get; set; }
}
