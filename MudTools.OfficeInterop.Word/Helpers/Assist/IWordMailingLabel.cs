//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示邮件标签的全局邮件标签首选项。
/// <para>注：使用 Application.MailingLabel 属性可返回 MailingLabel 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordMailingLabel : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    #region 邮件标签属性 (Mailing Label Properties)

    /// <summary>
    /// 获取或设置默认的邮件标签编号。
    /// </summary>
    string DefaultLabelName { get; set; }

    /// <summary>
    /// 获取或设置是否默认打印条形码。
    /// </summary>
    bool DefaultPrintBarCode { get; set; }

    /// <summary>
    /// 获取或设置默认激光打印机纸盒。
    /// </summary>
    WdPaperTray DefaultLaserTray { get; set; }

    /// <summary>
    /// 获取自定义标签对象。
    /// </summary>
    IWordCustomLabels? CustomLabels { get; }
    #endregion

    #region 邮件标签方法 (Mailing Label Methods)
    void LabelOptions();
    /// <summary>
    /// 创建一个邮件标签。
    /// </summary>
    /// <param name="name">标签名称。</param>
    /// <param name="address">地址。</param>
    /// <param name="autoText">自动图文集。</param>
    /// <param name="extractAddress">提取地址。</param>
    /// <param name="laserTray">激光托盘。</param>
    /// <param name="printEPostageLabel">使用互联网电子邮政供应商打印邮资。</param>
    /// <param name="vertical">格式化文本标签的垂直方向上。 用于亚洲语言邮件标签。</param>
    /// <returns>表示创建的邮件标签范围。</returns>
    IWordDocument? CreateNewDocument(string? name = null, string? address = null,
        string? autoText = null, bool? extractAddress = null,
        WdPaperTray laserTray = WdPaperTray.wdPrinterDefaultBin,
        bool? printEPostageLabel = null, bool? vertical = null);

    /// <summary>
    /// 打印邮件标签。
    /// </summary>
    /// <param name="name">标签名称。</param>
    /// <param name="address">地址。</param>
    /// <param name="extractAddress">提取地址。</param>
    /// <param name="laserTray">激光托盘。</param>
    void PrintOut(string? name, string? address, string? extractAddress,
         WdPaperTray laserTray = WdPaperTray.wdPrinterDefaultBin);

    /// <summary>
    /// 创建一个邮件标签，根据指定的标签ID。
    /// </summary>
    /// <param name="labelID">标签ID。</param>
    /// <param name="address">地址。</param>
    /// <param name="autoText">自动图文集。</param>
    /// <param name="extractAddress">提取地址。</param>
    /// <param name="laserTray">激光托盘。</param>
    /// <param name="printEPostageLabel">使用互联网电子邮政供应商打印邮资。</param>
    /// <param name="vertical">格式化文本标签的垂直方向上。用于亚洲语言邮件标签。</param>
    /// <returns>表示创建的邮件标签文档，如果创建失败则返回null。</returns>
    IWordDocument? CreateNewDocumentByID(string? labelID = null, string? address = null, string? autoText = null,
                                         bool? extractAddress = null, WdPaperTray? laserTray = null,
                                         bool? printEPostageLabel = null, bool? vertical = null);


    /// <summary>
    /// 根据指定的标签ID打印邮件标签。
    /// </summary>
    /// <param name="labelID">标签ID。</param>
    /// <param name="address">地址。</param>
    /// <param name="extractAddress">提取地址。</param>
    /// <param name="laserTray">激光托盘。</param>
    /// <param name="singleLabel">是否为单个标签。</param>
    /// <param name="row">标签行位置。</param>
    /// <param name="column">标签列位置。</param>
    /// <param name="printEPostageLabel">使用互联网电子邮政供应商打印邮资。</param>
    /// <param name="vertical">格式化文本标签的垂直方向上。用于亚洲语言邮件标签。</param>
    void PrintOutByID(string? labelID, string? address, string? extractAddress,
                     WdPaperTray laserTray, bool? singleLabel, object row, string? column,
                     bool? printEPostageLabel, bool? vertical);


    #endregion
}