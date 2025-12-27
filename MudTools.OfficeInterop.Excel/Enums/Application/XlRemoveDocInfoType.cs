//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要从文档信息中移除的信息类型
/// </summary>
public enum XlRemoveDocInfoType
{
    /// <summary>
    /// 从文档信息中移除批注
    /// </summary>
    xlRDIComments = 1,

    /// <summary>
    /// 从文档信息中移除个人信息
    /// </summary>
    xlRDIRemovePersonalInformation = 4,

    /// <summary>
    /// 从文档信息中移除电子邮件标头
    /// </summary>
    xlRDIEmailHeader = 5,

    /// <summary>
    /// 从文档信息中移除传送名单信息
    /// </summary>
    xlRDIRoutingSlip = 6,

    /// <summary>
    /// 从文档信息中移除发送供审阅信息
    /// </summary>
    xlRDISendForReview = 7,

    /// <summary>
    /// 从文档信息中移除文档属性
    /// </summary>
    xlRDIDocumentProperties = 8,

    /// <summary>
    /// 从文档信息中移除工作区数据
    /// </summary>
    xlRDIDocumentWorkspace = 10,

    /// <summary>
    /// 从文档信息中移除墨迹批注
    /// </summary>
    xlRDIInkAnnotations = 11,

    /// <summary>
    /// 从文档信息中移除方案批注
    /// </summary>
    xlRDIScenarioComments = 12,

    /// <summary>
    /// 从文档信息中移除发布信息数据
    /// </summary>
    xlRDIPublishInfo = 13,

    /// <summary>
    /// 从文档信息中移除服务器属性
    /// </summary>
    xlRDIDocumentServerProperties = 14,

    /// <summary>
    /// 从文档信息中移除文档管理策略数据
    /// </summary>
    xlRDIDocumentManagementPolicy = 15,

    /// <summary>
    /// 从文档信息中移除内容类型数据
    /// </summary>
    xlRDIContentType = 16,

    /// <summary>
    /// 从文档信息中移除定义名称批注
    /// </summary>
    xlRDIDefinedNameComments = 18,

    /// <summary>
    /// 从文档信息中移除非活动数据连接数据
    /// </summary>
    xlRDIInactiveDataConnections = 19,

    /// <summary>
    /// 从文档信息中移除打印机路径
    /// </summary>
    xlRDIPrinterPath = 20,

    /// <summary>
    /// 从文档信息中移除内联 Web 扩展
    /// </summary>
    xlRDIInlineWebExtensions = 21,

    /// <summary>
    /// 从文档信息中移除任务窗格 Web 扩展
    /// </summary>
    xlRDITaskpaneWebExtensions = 22,

    /// <summary>
    /// 从文档信息中移除 Excel 数据模型
    /// </summary>
    xlRDIExcelDataModel = 23,

    /// <summary>
    /// 移除所有文档信息
    /// </summary>
    xlRDIAll = 99
}