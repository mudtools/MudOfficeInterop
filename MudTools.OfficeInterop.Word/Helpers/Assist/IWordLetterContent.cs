//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示通过"信函向导"创建的信函内容。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLetterContent : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个只读的 LetterContent 对象，表示通过"信函向导"创建的指定信函内容。
    /// </summary>
    IWordLetterContent? Duplicate { get; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的日期格式。
    /// </summary>
    string DateFormat { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在"信函向导"创建的信函中包含页眉和页脚。
    /// </summary>
    bool IncludeHeaderFooter { get; set; }

    /// <summary>
    /// 获取或设置附加到"信函向导"创建文档的模板名称。
    /// </summary>
    string PageDesign { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的布局样式。
    /// </summary>
    WdLetterStyle LetterStyle { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为"信函向导"创建的信函预留预印信头空间。
    /// </summary>
    bool Letterhead { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函中预印信头的位置。
    /// </summary>
    WdLetterheadLocation LetterheadLocation { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函中预印信头预留的空间大小（以磅为单位）。
    /// </summary>
    float LetterheadSize { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的收件人姓名。
    /// </summary>
    string RecipientName { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的收件人地址。
    /// </summary>
    string RecipientAddress { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的称呼文本。
    /// </summary>
    string Salutation { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的称呼类型。
    /// </summary>
    WdSalutationType SalutationType { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的参考行（例如："回复："）。
    /// </summary>
    string RecipientReference { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的邮寄指示文本（例如："挂号信"）。
    /// </summary>
    string MailingInstructions { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的注意行文本。
    /// </summary>
    string AttentionLine { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的主题文本。
    /// </summary>
    string Subject { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的附件数量。
    /// </summary>
    int EnclosureNumber { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的抄送（CC）收件人。
    /// </summary>
    string CCList { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的寄信人地址。
    /// </summary>
    string ReturnAddress { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建信函的人员姓名。
    /// </summary>
    string SenderName { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建的信函的结尾文本（例如："此致敬礼"）。
    /// </summary>
    string Closing { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建信函的人员的公司名称。
    /// </summary>
    string SenderCompany { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建信函的人员的职位。
    /// </summary>
    string SenderJobTitle { get; set; }

    /// <summary>
    /// 获取或设置通过"信函向导"创建信函的人员的姓名缩写。
    /// </summary>
    string SenderInitials { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值与 Microsoft Word 中的"信函向导"相关联。在美国英语版本的 Word 中未使用。
    /// </summary>
    bool InfoBlock { get; set; }

    /// <summary>
    /// 获取或设置收件人代码。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    string RecipientCode { get; set; }

    /// <summary>
    /// 获取或设置收件人性别（如果已知）。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    WdSalutationGender RecipientGender { get; set; }

    /// <summary>
    /// 获取或设置简写地址。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    string ReturnAddressShortForm { get; set; }

    /// <summary>
    /// 获取或设置寄信人城市。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    string SenderCity { get; set; }

    /// <summary>
    /// 获取或设置寄信人代码。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    string SenderCode { get; set; }

    /// <summary>
    /// 获取或设置用于称呼的性别。在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    WdSalutationGender SenderGender { get; set; }

    /// <summary>
    /// 在美国英语版本的 Microsoft Word 中未使用。
    /// </summary>
    string SenderReference { get; set; }

}