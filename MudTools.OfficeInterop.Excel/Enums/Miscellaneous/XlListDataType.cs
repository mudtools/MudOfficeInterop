//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定列表数据类型的枚举，用于定义Excel工作表中列表列的数据类型
/// </summary>
public enum XlListDataType
{
    /// <summary>
    /// 无特定数据类型
    /// </summary>
    xlListDataTypeNone,

    /// <summary>
    /// 文本数据类型
    /// </summary>
    xlListDataTypeText,

    /// <summary>
    /// 多行文本数据类型
    /// </summary>
    xlListDataTypeMultiLineText,

    /// <summary>
    /// 数字数据类型
    /// </summary>
    xlListDataTypeNumber,

    /// <summary>
    /// 货币数据类型
    /// </summary>
    xlListDataTypeCurrency,

    /// <summary>
    /// 日期时间数据类型
    /// </summary>
    xlListDataTypeDateTime,

    /// <summary>
    /// 选择数据类型（单选）
    /// </summary>
    xlListDataTypeChoice,

    /// <summary>
    /// 选择数据类型（多选）
    /// </summary>
    xlListDataTypeChoiceMulti,

    /// <summary>
    /// 列表查找数据类型
    /// </summary>
    xlListDataTypeListLookup,

    /// <summary>
    /// 复选框数据类型
    /// </summary>
    xlListDataTypeCheckbox,

    /// <summary>
    /// 超链接数据类型
    /// </summary>
    xlListDataTypeHyperLink,

    /// <summary>
    /// 计数器数据类型
    /// </summary>
    xlListDataTypeCounter,

    /// <summary>
    /// 多行富文本数据类型
    /// </summary>
    xlListDataTypeMultiLineRichText
}