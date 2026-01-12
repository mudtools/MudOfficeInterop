//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定范围值的数据类型
/// </summary>
public enum XlRangeValueDataType
{
    /// <summary>
    /// 默认值。如果指定的 Range 对象为空，则返回 Empty 值（使用 IsEmpty 函数测试这种情况）。如果 Range 对象包含多个单元格，则返回值数组（使用 IsArray 函数测试这种情况）
    /// </summary>
    xlRangeValueDefault = 10,

    /// <summary>
    /// 以 XML 电子表格格式返回指定 Range 对象的值、格式、公式和名称
    /// </summary>
    xlRangeValueXMLSpreadsheet,

    /// <summary>
    /// 以 XML 格式返回指定 Range 对象的记录集表示形式
    /// </summary>
    xlRangeValueMSPersistXML
}