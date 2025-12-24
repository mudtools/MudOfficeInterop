//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定要从文档中移除的信息类型
/// </summary>
public enum WdRemoveDocInfoType
{
    /// <summary>
    /// 移除文档批注
    /// </summary>
    wdRDIComments = 1,

    /// <summary>
    /// 移除修订标记
    /// </summary>
    wdRDIRevisions = 2,

    /// <summary>
    /// 移除文档版本信息
    /// </summary>
    wdRDIVersions = 3,

    /// <summary>
    /// 移除个人信息
    /// </summary>
    wdRDIRemovePersonalInformation = 4,

    /// <summary>
    /// 移除电子邮件标头信息
    /// </summary>
    wdRDIEmailHeader = 5,

    /// <summary>
    /// 移除传送名单信息
    /// </summary>
    wdRDIRoutingSlip = 6,

    /// <summary>
    /// 移除发送供审阅时存储的信息
    /// </summary>
    wdRDISendForReview = 7,

    /// <summary>
    /// 移除文档属性
    /// </summary>
    wdRDIDocumentProperties = 8,

    /// <summary>
    /// 移除模板信息
    /// </summary>
    wdRDITemplate = 9,

    /// <summary>
    /// 移除文档工作区信息
    /// </summary>
    wdRDIDocumentWorkspace = 10,

    /// <summary>
    /// 移除墨迹批注
    /// </summary>
    wdRDIInkAnnotations = 11,

    /// <summary>
    /// 移除文档服务器属性
    /// </summary>
    wdRDIDocumentServerProperties = 14,

    /// <summary>
    /// 移除文档管理策略信息
    /// </summary>
    wdRDIDocumentManagementPolicy = 15,

    /// <summary>
    /// 移除内容类型信息
    /// </summary>
    wdRDIContentType = 16,

    /// <summary>
    /// 移除任务窗格 Web 扩展信息
    /// </summary>
    wdRDITaskpaneWebExtensions = 17,

    /// <summary>
    /// 移除所有文档信息
    /// </summary>
    wdRDIAll = 99
}