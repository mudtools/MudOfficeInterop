//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定Excel工作簿中外部数据连接的类型
/// </summary>
public enum XlConnectionType
{
    /// <summary>
    /// OLE DB 连接类型
    /// </summary>
    xlConnectionTypeOLEDB = 1,
    /// <summary>
    /// ODBC 连接类型
    /// </summary>
    xlConnectionTypeODBC,
    /// <summary>
    /// XML 映射连接类型
    /// </summary>
    xlConnectionTypeXMLMAP,
    /// <summary>
    /// 文本文件连接类型
    /// </summary>
    xlConnectionTypeTEXT,
    /// <summary>
    /// Web 数据连接类型
    /// </summary>
    xlConnectionTypeWEB,
    /// <summary>
    /// 数据馈送连接类型
    /// </summary>
    xlConnectionTypeDATAFEED,
    /// <summary>
    /// 模型连接类型
    /// </summary>
    xlConnectionTypeMODEL,
    /// <summary>
    /// 工作表连接类型
    /// </summary>
    xlConnectionTypeWORKSHEET,
    /// <summary>
    /// 无数据源连接类型
    /// </summary>
    xlConnectionTypeNOSOURCE
}