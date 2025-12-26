//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 文档中信封功能的托管包装接口。
/// 每个 Word 文档仅包含一个信封对象，可通过 Document.Envelope 属性获取 [[1]]。
/// 该对象始终可用，无论信封是否已插入到文档中 [[4]]。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordEnvelope : IOfficeObject<IWordEnvelope>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordEnvelope"/> 对象的父对象。
    /// </summary>
    /// <remarks>父对象通常是 TextColumns 集合。</remarks>
    object? Parent { get; }

    #region 信封属性
    /// <summary>
    /// 获取或设置收件人地址 [[3]]。
    /// </summary>
    IWordRange? Address { get; }

    /// <summary>
    /// 获取或设置发件人（回邮）地址 [[3]]。
    /// </summary>
    IWordRange? ReturnAddress { get; }

    /// <summary>
    /// 获取或设置信封纸张来源（打印机托盘）[[3]]。
    /// </summary>
    WdPaperTray FeedSource { get; set; }

    /// <summary>
    /// 获取或设置收件人地址距离信封左边界的距离（以磅为单位）[[3]]。
    /// </summary>
    float AddressFromLeft { get; set; }

    /// <summary>
    /// 获取或设置收件人地址距离信封上边界的距离（以磅为单位）[[3]]。
    /// </summary>
    float AddressFromTop { get; set; }

    /// <summary>
    /// 获取或设置回邮地址距离信封左边界的距离（以磅为单位）[[3]]。
    /// </summary>
    float ReturnAddressFromLeft { get; set; }

    /// <summary>
    /// 获取或设置回邮地址距离信封上边界的距离（以磅为单位）[[3]]。
    /// </summary>
    float ReturnAddressFromTop { get; set; }

    /// <summary>
    /// 获取或设置信封的默认宽度（以磅为单位）。
    /// </summary>
    float DefaultWidth { get; set; }

    /// <summary>
    /// 获取或设置信封的默认高度（以磅为单位）。
    /// </summary>
    float DefaultHeight { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否默认打印条形码。
    /// </summary>
    bool DefaultPrintBarCode { get; set; }

    /// <summary>
    /// 获取或设置信封的默认尺寸。
    /// </summary>
    string DefaultSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示信封是否垂直放置。
    /// </summary>
    bool Vertical { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示信封是否默认正面朝上。
    /// </summary>
    bool DefaultFaceUp { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否默认省略回邮地址。
    /// </summary>
    bool DefaultOmitReturnAddress { get; set; }

    /// <summary>
    /// 获取或设置信封的默认方向。
    /// </summary>
    WdEnvelopeOrientation DefaultOrientation { get; set; }

    /// <summary>
    /// 获取回邮地址的样式设置。
    /// </summary>
    IWordStyle? ReturnAddressStyle { get; }

    /// <summary>
    /// 获取收件人地址的样式设置。
    /// </summary>
    IWordStyle? AddressStyle { get; }

    /// <summary>
    /// 获取或设置收件人姓名距离信封左边界的距离（以磅为单位）。
    /// </summary>
    float RecipientNamefromLeft { get; set; }

    /// <summary>
    /// 获取或设置收件人姓名距离信封上边界的距离（以磅为单位）。
    /// </summary>
    float RecipientNamefromTop { get; set; }

    /// <summary>
    /// 获取或设置收件人邮政编码距离信封左边界的距离（以磅为单位）。
    /// </summary>
    float RecipientPostalfromLeft { get; set; }

    /// <summary>
    /// 获取或设置收件人邮政编码距离信封上边界的距离（以磅为单位）。
    /// </summary>
    float RecipientPostalfromTop { get; set; }

    /// <summary>
    /// 获取或设置发件人姓名距离信封左边界的距离（以磅为单位）。
    /// </summary>
    float SenderNamefromLeft { get; set; }

    /// <summary>
    /// 获取或设置发件人姓名距离信封上边界的距离（以磅为单位）。
    /// </summary>
    float SenderNamefromTop { get; set; }

    /// <summary>
    /// 获取或设置发件人邮政编码距离信封左边界的距离（以磅为单位）。
    /// </summary>
    float SenderPostalfromLeft { get; set; }

    /// <summary>
    /// 获取或设置发件人邮政编码距离信封上边界的距离（以磅为单位）。
    /// </summary>
    float SenderPostalfromTop { get; set; }
    #endregion

    #region 信封操作方法

    /// <summary>
    /// 显示信封选项对话框，允许用户设置信封的各种选项 [[7]]。
    /// </summary>
    void Options();

    /// <summary>
    /// 在指定文档的开头以单独一节的形式插入一个信封 [[7]]。
    /// </summary>
    /// <param name="address">收件人地址。如果指定了 AutoText 条目，则此参数将被忽略。</param>
    /// <param name="returnAddress">发件人地址。</param>
    /// <param name="autoText">用于地址的自动图文集条目名称。</param>
    /// <param name="omitReturnAddress">如果为 true，则省略回邮地址。</param>
    /// <param name="printBarcode">如果为 true，则打印邮政条形码。</param>
    /// <param name="printFIMA">如果为 true，则打印 FIMA 条。</param>
    /// <param name="size">信封尺寸 [[3]]。</param>
    /// <param name="feedSource">信封的纸张来源 [[7]]。</param>
    void Insert(
        string? address = null,
        string? returnAddress = null,
        string? autoText = null,
        bool omitReturnAddress = false,
        bool printBarcode = false,
        bool printFIMA = false,
        string? size = null,
        int? feedSource = null);

    /// <summary>
    /// 打印信封，但不将信封添加到活动文档中 [[7]]。
    /// </summary>
    /// <param name="address">收件人地址。如果指定了 AutoText 条目，则此参数将被忽略。</param>
    /// <param name="returnAddress">发件人地址。</param>
    /// <param name="autoText">用于地址的自动图文集条目名称。</param>
    /// <param name="omitReturnAddress">如果为 true，则省略回邮地址。</param>
    /// <param name="printBarcode">如果为 true，则打印邮政条形码。</param>
    /// <param name="printFIMA">如果为 true，则打印 FIMA 条。</param>
    /// <param name="size">信封尺寸。</param>
    /// <param name="feedSource">信封的纸张来源。</param>
    void PrintOut(
        string? address = null,
        string? returnAddress = null,
        string? autoText = null,
        bool omitReturnAddress = false,
        bool printBarcode = false,
        bool printFIMA = false,
        string? size = null,
        int? feedSource = null);

    /// <summary>
    /// 更新文档中的信封内容以匹配当前信封对象的属性 [[4]]。
    /// </summary>
    void UpdateDocument();

    #endregion
}