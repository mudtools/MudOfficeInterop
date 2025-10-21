//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定邮件合并操作的数据源子类型
/// </summary>
public enum WdMergeSubType
{
    /// <summary>
    /// 其他或未知数据源类型
    /// </summary>
    wdMergeSubTypeOther,

    /// <summary>
    /// Microsoft Access 数据库
    /// </summary>
    wdMergeSubTypeAccess,

    /// <summary>
    /// Office Address List (OAL)
    /// </summary>
    wdMergeSubTypeOAL,

    /// <summary>
    /// OLE DB Word 数据源
    /// </summary>
    wdMergeSubTypeOLEDBWord,

    /// <summary>
    /// Microsoft Works 数据源
    /// </summary>
    wdMergeSubTypeWorks,

    /// <summary>
    /// OLE DB 文本数据源
    /// </summary>
    wdMergeSubTypeOLEDBText,

    /// <summary>
    /// Microsoft Outlook 联系人列表
    /// </summary>
    wdMergeSubTypeOutlook,

    /// <summary>
    /// Microsoft Word 文档作为数据源
    /// </summary>
    wdMergeSubTypeWord,

    /// <summary>
    /// Microsoft Word 2000 文档作为数据源
    /// </summary>
    wdMergeSubTypeWord2000
}