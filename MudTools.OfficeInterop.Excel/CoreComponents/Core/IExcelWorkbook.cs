//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Vbe;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Workbook 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Workbook 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel", NoneConstructor = true)]
public interface IExcelWorkbook : IDisposable
{

    /// <summary>
    /// 获取表示Excel应用程序的Application对象。当不带对象限定符使用时，此属性返回表示Excel应用程序的Application对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }


    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否可以在工作表中使用标签。默认值为False。
    /// </summary>
    bool AcceptLabelsInFormulas { get; set; }

    /// <summary>
    /// 激活与工作簿关联的第一个窗口。这不会运行可能附加到工作簿的任何Auto_Activate或Auto_Deactivate宏。
    /// </summary>
    void Activate();

    /// <summary>
    /// 获取表示活动图表（嵌入图表或图表工作表）的Chart对象。当没有活动图表时，此属性返回null。
    /// </summary>
    IExcelChart ActiveChart { get; }

    /// <summary>
    /// 获取活动工作簿或指定窗口或工作簿中的活动工作表（位于顶部的）。如果没有活动工作表，则返回null。
    /// </summary>
    object ActiveSheet { get; }

    /// <summary>
    /// 获取或设置作者的名称。
    /// </summary>
    string Author { get; set; }

    /// <summary>
    /// 获取或设置共享工作簿自动更新之间的分钟数。
    /// </summary>
    int AutoUpdateFrequency { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每当工作簿自动更新时，当前对共享工作簿的更改是否发布给其他用户。默认为True。
    /// </summary>
    bool AutoUpdateSaveChanges { get; set; }

    /// <summary>
    /// 获取或设置共享工作簿变更历史记录显示的天数。
    /// </summary>
    int ChangeHistoryDuration { get; set; }

    /// <summary>
    /// 获取表示工作簿所有内置文档属性的DocumentProperties集合。
    /// </summary>
    object BuiltinDocumentProperties { get; }

    /// <summary>
    /// 更改工作簿的访问权限。这可能需要从磁盘加载更新版本。
    /// </summary>
    /// <param name="mode">必需。指定新的访问模式。</param>
    /// <param name="writePassword">可选。如果文件是写保护的且模式为xlReadWrite，则指定写保护密码。</param>
    /// <param name="notify">可选。True（或省略）表示如果文件无法立即访问则通知用户。</param>
    void ChangeFileAccess(XlFileAccess mode, string? writePassword = null, bool? notify = null);

    /// <summary>
    /// 将链接从一个文档更改为另一个文档。
    /// </summary>
    /// <param name="name">必需。要更改的Excel或DDE/OLE链接的名称，由LinkSources方法返回。</param>
    /// <param name="newName">必需。链接的新名称。</param>
    /// <param name="type">可选。链接类型。</param>
    void ChangeLink(string name, string newName, XlLinkType type = XlLinkType.xlLinkTypeExcelLinks);

    /// <summary>
    /// 获取表示工作簿中所有图表工作表的Sheets集合。
    /// </summary>
    IExcelSheets? Charts { get; }

    /// <summary>
    /// 关闭工作簿。
    /// </summary>
    /// <param name="saveChanges">可选。如果没有对工作簿的更改，则忽略此参数。如果有更改且工作簿出现在其他打开的窗口中，也忽略此参数。如果有更改但工作簿未出现在任何其他打开的窗口中，此参数指定是否保存更改。</param>
    /// <param name="filename">可选。将更改保存到此文件名下。</param>
    /// <param name="routeWorkbook">可选。如果工作簿不需要路由到下一个收件人，则忽略此参数。否则，Microsoft Excel根据以下值路由工作簿：True表示发送给下一个收件人，False表示不发送，省略则显示对话框询问用户是否发送。</param>
    void Close(bool? saveChanges = null, string? filename = null, bool? routeWorkbook = null);

    /// <summary>
    /// 获取对象的代码名称。在设计时，可以通过更改此值来更改对象的代码名称。运行时无法通过编程更改此属性。
    /// </summary>
    string CodeName { get; }

    /// <summary>
    /// 获取或设置工作簿调色板中的颜色。调色板有56个条目，每个条目由RGB值表示。
    /// </summary>
    /// <param name="index">可选。颜色编号（从1到56）。如果未指定此参数，则返回包含调色板所有56种颜色的数组。</param>
    object Colors { get; set; }

    /// <summary>
    /// 获取表示Excel命令栏的CommandBars对象。
    /// </summary>
    IOfficeCommandBars? CommandBars { get; }

    /// <summary>
    /// 获取表示所有批注的Comments集合。
    /// </summary>
    string Comments { get; set; }

    /// <summary>
    /// 获取或设置每当更新共享工作簿时解决冲突的方式。
    /// </summary>
    XlSaveConflictResolution ConflictResolution { get; set; }

    /// <summary>
    /// 获取表示指定OLE对象容器应用程序的对象。
    /// </summary>
    object Container { get; }

    /// <summary>
    /// 获取一个布尔值，表示保存此文件时是否创建备份文件。
    /// </summary>
    bool CreateBackup { get; }

    /// <summary>
    /// 获取或设置表示工作簿所有自定义文档属性的DocumentProperties集合。
    /// </summary>
    object CustomDocumentProperties { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示工作簿是否使用1904日期系统。
    /// </summary>
    bool Date1904 { get; set; }

    /// <summary>
    /// 从工作簿中删除自定义数字格式。
    /// </summary>
    /// <param name="numberFormat">必需。要删除的数字格式名称。</param>
    void DeleteNumberFormat(string numberFormat);

    /// <summary>
    /// 获取或设置形状的显示方式。
    /// </summary>
    XlDisplayDrawingObjects DisplayDrawingObjects { get; set; }

    /// <summary>
    /// 为当前用户分配对作为共享列表打开的工作簿的独占访问权限。
    /// </summary>
    /// <returns>操作结果。</returns>
    bool ExclusiveAccess();

    /// <summary>
    /// 获取工作簿的文件格式和/或类型。
    /// </summary>
    XlFileFormat FileFormat { get; }

    /// <summary>
    /// 获取对象的完整名称，包括其在磁盘上的路径。
    /// </summary>
    string FullName { get; }

    /// <summary>
    /// 获取一个布尔值，表示工作簿是否有保护密码。
    /// </summary>
    bool HasPassword { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示工作簿是否有路由单。
    /// </summary>
    bool HasRoutingSlip { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示工作簿是否作为加载项运行。
    /// </summary>
    bool IsAddin { get; set; }

    /// <summary>
    /// 获取链接的日期和更新状态。
    /// </summary>
    /// <param name="name">可选。链接名称。</param>
    /// <param name="linkInfo">必需。要返回的信息类型。</param>
    /// <param name="type">可选。要返回的链接类型。</param>
    /// <param name="editionRef">可选。如果链接是版本，此参数将版本引用指定为R1C1样式的字符串。如果工作簿中有多个具有相同名称的发布者或订阅者，则需要此参数。</param>
    /// <returns>链接信息。</returns>
    object? LinkInfo(string name, XlLinkInfo linkInfo, XlLinkInfoType? type = null, string? editionRef = null);

    /// <summary>
    /// 获取工作簿中的链接数组。数组中的名称是链接文档、版本或DDE或OLE服务器的名称。如果没有链接，则返回空值。
    /// </summary>
    /// <param name="type">可选。要返回的链接类型。</param>
    /// <returns>链接源数组。</returns>
    object? LinkSources(XlLinkType? type = null);

    /// <summary>
    /// 将另一个工作簿中的更改合并到打开的工作簿中。
    /// </summary>
    /// <param name="filename">必需。包含要合并到打开工作簿中的更改的工作簿的文件名。</param>
    void MergeWorkbook(string? filename);

    /// <summary>
    /// 获取一个布尔值，表示工作簿是否作为共享列表打开。
    /// </summary>
    bool MultiUserEditing { get; }

    /// <summary>
    /// 获取工作簿的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取表示工作簿中所有名称（包括工作表特定名称）的Names集合。
    /// </summary>
    IExcelNames? Names { get; }

    /// <summary>
    /// 为指定窗口创建新窗口或副本。
    /// </summary>
    /// <returns>新创建的Window对象。</returns>
    IExcelWindow? NewWindow();

    /// <summary>
    /// 打开链接的支持文档。
    /// </summary>
    /// <param name="name">必需。要打开的Excel或DDE/OLE链接的名称，由LinkSources方法返回。</param>
    /// <param name="readOnly">可选。True表示以只读方式打开文档。默认值为False。</param>
    /// <param name="type">可选。链接类型。</param>
    void OpenLinks(string name, bool? readOnly = null, XlLinkType? type = null);

    /// <summary>
    /// 获取应用程序的完整路径，不包括最终分隔符和应用程序名称。
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示列表的筛选和排序设置是否包含在用户对共享工作簿的个人视图中。
    /// </summary>
    bool PersonalViewListSettings { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示打印设置是否包含在用户对共享工作簿的个人视图中。
    /// </summary>
    bool PersonalViewPrintSettings { get; set; }

    /// <summary>
    /// 获取表示工作簿中所有数据透视表缓存的PivotCaches集合。
    /// </summary>
    /// <returns>数据透视表缓存集合。</returns>
    IExcelPivotCaches PivotCaches();

    /// <summary>
    /// 将指定的工作簿发布到公共文件夹。此方法仅适用于连接到Microsoft Exchange服务器的Microsoft Exchange客户端。
    /// </summary>
    /// <param name="destName">可选。此参数被忽略。Post方法会提示用户指定工作簿的目标位置。</param>
    void Post(string? destName = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示此工作簿中的计算是否仅使用数字显示时的精度。
    /// </summary>
    bool PrecisionAsDisplayed { get; set; }

    /// <summary>
    /// 显示对象的打印预览。
    /// </summary>
    /// <param name="enableChanges">启用对对象的更改。</param>
    void PrintPreview(bool? enableChanges = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示是否将工作簿保护用于共享。
    /// </summary>
    /// <param name="filename">可选。指示保存文件名的字符串。可以包含完整路径；如果不包含，Microsoft Excel将文件保存在当前文件夹中。</param>
    /// <param name="password">可选。区分大小写的字符串，表示给文件的保护密码。长度不应超过15个字符。</param>
    /// <param name="writeResPassword">可选。指示此文件的写保留密码的字符串。如果文件保存时使用密码，并且打开文件时未提供密码，则文件以只读方式打开。</param>
    /// <param name="readOnlyRecommended">可选。True表示打开文件时显示消息，建议以只读方式打开文件。</param>
    /// <param name="createBackup">可选。True表示创建备份文件。</param>
    /// <param name="sharingPassword">可选。用于保护文件共享的密码字符串。</param>
    void ProtectSharing(string? filename = null, string? password = null,
                        string? writeResPassword = null, bool? readOnlyRecommended = null,
                        bool? createBackup = null, string? sharingPassword = null);

    /// <summary>
    /// 获取一个布尔值，表示工作簿中的工作表顺序是否受保护。
    /// </summary>
    bool ProtectStructure { get; }

    /// <summary>
    /// 获取一个布尔值，表示工作簿的窗口是否受保护。
    /// </summary>
    bool ProtectWindows { get; }

    /// <summary>
    /// 获取一个布尔值，表示对象是否以只读方式打开。
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 刷新工作簿中的所有外部数据范围和透视表报告。
    /// </summary>
    void RefreshAll();

    /// <summary>
    /// 从共享工作簿中断开指定用户的连接。
    /// </summary>
    /// <param name="index">必需。用户索引。</param>
    void RemoveUser(int index);

    /// <summary>
    /// 获取工作簿作为共享列表打开时已保存的次数。如果工作簿以独占模式打开，此属性返回0。
    /// </summary>
    int RevisionNumber { get; }

    /// <summary>
    /// 使用工作簿的当前路由单路由工作簿。
    /// </summary>
    void Route();

    /// <summary>
    /// 获取一个布尔值，表示工作簿是否已路由到下一个收件人。如果工作簿需要路由，则返回False。
    /// </summary>
    bool Routed { get; }

    /// <summary>
    /// 运行附加到工作簿的Auto_Open、Auto_Close、Auto_Activate或Auto_Deactivate宏。此方法包含用于向后兼容性。
    /// </summary>
    /// <param name="which">必需。指定要运行的宏类型。</param>
    void RunAutoMacros(XlRunAutoMacro which);

    /// <summary>
    /// 保存对工作簿的更改。
    /// </summary>
    void Save();

    /// <summary>
    /// 将工作簿的副本保存到文件，但不修改内存中打开的工作簿。
    /// </summary>
    /// <param name="filename">必需。指定副本的文件名。</param>
    void SaveCopyAs(string? filename = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示自上次保存以来是否未对工作簿进行更改。
    /// </summary>
    bool Saved { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示Microsoft Excel是否将外部链接值与工作簿一起保存。
    /// </summary>
    bool SaveLinkValues { get; set; }

    /// <summary>
    /// 使用已安装的邮件系统发送工作簿。
    /// </summary>
    /// <param name="recipients">必需。指定收件人姓名的字符串，或包含多个收件人的文本字符串数组。至少必须指定一个收件人，所有收件人都添加为"收件人"。</param>
    /// <param name="subject">可选。消息的主题。如果省略此参数，则使用文档名称。</param>
    /// <param name="returnReceipt">可选。True表示请求回执。False表示不请求回执。默认值为False。</param>
    void SendMail(string? recipients, string? subject = null, bool? returnReceipt = null);

    /// <summary>
    /// 设置每当DDE链接更新时运行的过程的名称。
    /// </summary>
    /// <param name="name">必需。DDE/OLE链接的名称，由LinkSources方法返回。</param>
    /// <param name="procedure">必需。链接更新时要运行的过程的名称。可以是Excel 4.0宏或Visual Basic过程。将此参数设置为空字符串表示链接更新时不运行任何过程。</param>
    void SetLinkOnData(string name, string? procedure = null);

    /// <summary>
    /// 获取表示工作簿中所有工作表的Sheets集合。
    /// </summary>
    IExcelSheets? Sheets { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示冲突历史记录工作表在作为共享列表打开的工作簿中是否可见。
    /// </summary>
    bool ShowConflictHistory { get; set; }

    /// <summary>
    /// 获取表示工作簿中所有样式的Styles集合。
    /// </summary>
    IExcelStyles? Styles { get; }

    /// <summary>
    /// 取消对工作表或工作簿的保护。如果工作表或工作簿未受保护，此方法无效。
    /// </summary>
    /// <param name="password">可选。用于取消保护工作表或工作簿的区分大小写的密码字符串。如果工作表或工作簿未使用密码保护，则忽略此参数。如果省略此参数且工作表受密码保护，则会提示输入密码。如果省略此参数且工作簿受密码保护，则方法将失败。</param>
    void Unprotect(string? password = null);

    /// <summary>
    /// 关闭共享保护并保存工作簿。
    /// </summary>
    /// <param name="sharingPassword">可选。工作簿密码。</param>
    void UnprotectSharing(string? sharingPassword = null);

    /// <summary>
    /// 从保存的磁盘版本更新只读工作簿（如果磁盘版本比加载到内存中的工作簿副本更新）。如果自工作簿加载以来磁盘副本未更改，则不会重新加载内存中的工作簿副本。
    /// </summary>
    void UpdateFromFile();

    /// <summary>
    /// 更新Excel、DDE或OLE链接。
    /// </summary>
    /// <param name="name">可选。要更新的Excel或DDE/OLE链接的名称，由LinkSources方法返回。</param>
    /// <param name="type">可选。链接类型。</param>
    void UpdateLink(string? name = null, XlLinkType? type = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示Microsoft Excel是否更新工作簿中的远程引用。
    /// </summary>
    bool UpdateRemoteReferences { get; set; }

    /// <summary>
    /// 获取一个二维数组，提供有关以共享列表打开工作簿的每个用户的信息。
    /// </summary>
    object UserStatus { get; }

    /// <summary>
    /// 获取表示工作簿所有自定义视图的CustomViews集合。
    /// </summary>
    IExcelCustomViews? CustomViews { get; }

    /// <summary>
    /// 获取表示工作簿中所有窗口的Windows集合。
    /// </summary>
    IExcelWindows? Windows { get; }

    /// <summary>
    /// 获取表示工作簿中所有工作表的Sheets集合。
    /// </summary>
    IExcelSheets? Worksheets { get; }

    /// <summary>
    /// 获取一个布尔值，表示工作簿是否为写保留。
    /// </summary>
    bool WriteReserved { get; }

    /// <summary>
    /// 获取当前对工作簿具有写入权限的用户名称。
    /// </summary>
    string WriteReservedBy { get; }

    /// <summary>
    /// 获取表示工作簿中所有Excel 4.0国际宏工作表的Sheets集合。
    /// </summary>
    IExcelSheets? Excel4IntlMacroSheets { get; }

    /// <summary>
    /// 获取表示工作簿中所有Excel 4.0宏工作表的Sheets集合。
    /// </summary>
    IExcelSheets? Excel4MacroSheets { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示当工作簿另存为模板时是否删除外部数据引用。
    /// </summary>
    bool TemplateRemoveExtData { get; set; }

    /// <summary>
    /// 控制如何在共享工作簿中显示更改。
    /// </summary>
    /// <param name="when">可选。要显示的更改。可以是以下XlHighlightChangesTime常量之一：xlSinceMyLastSave、xlAllChanges或xlNotYetReviewed。</param>
    /// <param name="who">可选。要显示其更改的用户。可以是"Everyone"、"Everyone but Me"或共享工作簿的用户之一的名字。</param>
    /// <param name="where">可选。指定要检查更改的区域的A1样式范围引用。</param>
    void HighlightChangesOptions(XlHighlightChangesTime? when = null, string? who = null, object? where = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示是否在屏幕上高亮显示共享工作簿的更改。
    /// </summary>
    bool HighlightChangesOnScreen { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否为共享工作簿启用更改跟踪。
    /// </summary>
    bool KeepChangeHistory { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否在新工作表中显示对共享工作簿的更改。
    /// </summary>
    bool ListChangesOnNewSheet { get; set; }

    /// <summary>
    /// 从工作簿的更改日志中删除条目。
    /// </summary>
    /// <param name="days">必需。更改日志中要保留的更改天数。</param>
    /// <param name="sharingPassword">可选。取消共享保护工作簿的密码。如果工作簿受共享密码保护且省略此参数，则会提示用户输入密码。</param>
    void PurgeChangeHistoryNow(int days, string? sharingPassword = null);

    /// <summary>
    /// 接受指定共享工作簿中的所有更改。
    /// </summary>
    /// <param name="when">可选。指定何时接受所有更改。</param>
    /// <param name="who">可选。指定由谁接受所有更改。</param>
    /// <param name="where">可选。指定在何处接受所有更改。</param>
    void AcceptAllChanges(XlHighlightChangesTime? when = null, string? who = null, object? where = null);

    /// <summary>
    /// 拒绝指定共享工作簿中的所有更改。
    /// </summary>
    /// <param name="when">可选。指定何时拒绝所有更改。</param>
    /// <param name="who">可选。指定由谁拒绝所有更改。</param>
    /// <param name="where">可选。指定在何处拒绝所有更改。</param>
    void RejectAllChanges(XlHighlightChangesTime? when = null, string? who = null, object? where = null);

    /// <summary>
    /// 将颜色调色板重置为默认颜色。
    /// </summary>
    void ResetColors();


    //IVBProject VBProject { get; }

    /// <summary>
    /// 如果已下载，则显示缓存的文档。否则，此方法将解析超链接，下载目标文档，并在相应的应用程序中显示文档。
    /// </summary>
    /// <param name="address">必需。目标文档的地址。</param>
    /// <param name="subAddress">可选。目标文档中的位置。默认值为空字符串。</param>
    /// <param name="newWindow">可选。True表示在新窗口中显示目标应用程序。默认值为False。</param>
    /// <param name="addHistory">可选。未使用。保留供将来使用。</param>
    /// <param name="extraInfo">可选。指定HTTP用于解析超链接的附加信息的字符串或字节数组。</param>
    /// <param name="method">可选。指定附加ExtraInfo的方式。</param>
    /// <param name="headerInfo">可选。指定HTTP请求的标头信息的字符串。默认值为空字符串。</param>
    void FollowHyperlink(string address, object subAddress = null, object newWindow = null, object addHistory = null, object extraInfo = null, object method = null, object headerInfo = null);

    /// <summary>
    /// 将快捷方式添加到工作簿或收藏夹的超链接。
    /// </summary>
    void AddToFavorites();

    /// <summary>
    /// 获取一个布尔值，表示指定的工作簿是否正在就地编辑。如果工作簿已在Excel中打开进行编辑，则返回False。
    /// </summary>
    bool IsInplace { get; }

    /// <summary>
    /// 显示指定工作簿的预览，就像它另存为网页一样。
    /// </summary>
    void WebPagePreview();

    /// <summary>
    /// 获取PublishObjects集合，表示工作簿中显示在服务器上的发布对象。
    /// </summary>
    PublishObjects PublishObjects { get; }

    /// <summary>
    /// 获取WebOptions集合，包含将文档另存为网页或打开网页时使用的工作簿级属性。
    /// </summary>
    WebOptions WebOptions { get; }

    /// <summary>
    /// 使用指定的文档编码基于HTML文档重新加载工作簿。
    /// </summary>
    /// <param name="encoding">必需。要应用于工作簿的编码。</param>
    void ReloadAs(MsoEncoding encoding);

    /// <summary>
    /// 获取一个布尔值，表示电子邮件撰写标题和信封工具栏是否都可见。
    /// </summary>
    bool EnvelopeVisible { get; set; }

    /// <summary>
    /// 获取一个数字，其最右边四位是次要计算引擎版本号，其他数字是Excel的主版本号。
    /// </summary>
    int CalculationVersion { get; }

    /// <summary>
    /// 获取一个布尔值，表示工作簿的Visual Basic for Applications项目是否已进行数字签名。
    /// </summary>
    bool VBASigned { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否可以显示数据透视表字段列表。默认为True。
    /// </summary>
    bool ShowPivotTableFieldList { get; set; }

    /// <summary>
    /// 获取或设置一个XlUpdateLinks常量，指示工作簿用于更新嵌入OLE链接的设置。
    /// </summary>
    XlUpdateLinks UpdateLinks { get; set; }

    /// <summary>
    /// 将链接到其他Excel源或OLE源的公式转换为值。
    /// </summary>
    /// <param name="name">必需。链接的名称。</param>
    /// <param name="type">必需。链接类型。</param>
    void BreakLink(string name, XlLinkType type);

    /// <summary>
    /// 将更改保存到其他文件。
    /// </summary>
    /// <param name="filename">可选。指示要保存的文件名的字符串。可以包含完整路径；如果不包含，Microsoft Excel将文件保存在当前文件夹中。</param>
    /// <param name="fileFormat">可选。保存文件时使用的文件格式。</param>
    /// <param name="password">可选。区分大小写的字符串（不超过15个字符），表示给文件的保护密码。</param>
    /// <param name="writeResPassword">可选。指示此文件写保留密码的字符串。如果文件保存时使用密码，并且打开文件时未提供密码，则文件以只读方式打开。</param>
    /// <param name="readOnlyRecommended">可选。True表示打开文件时显示消息，建议以只读方式打开文件。</param>
    /// <param name="createBackup">可选。True表示创建备份文件。</param>
    /// <param name="accessMode">可选。保存访问模式。</param>
    /// <param name="conflictResolution">可选。冲突解决方式。</param>
    /// <param name="addToMru">可选。True表示将此工作簿添加到最近使用的文件列表。默认值为False。</param>
    /// <param name="textCodepage">可选。不用于美式英语Excel。</param>
    /// <param name="textVisualLayout">可选。不用于美式英语Excel。</param>
    /// <param name="local">可选。True表示根据Excel的语言（包括控制面板设置）保存文件。False（默认）表示根据VBA的语言（通常是美式英语）保存文件。</param>
    void SaveAs(object filename = null, object fileFormat = null, object password = null, object writeResPassword = null, object readOnlyRecommended = null, object createBackup = null, XlSaveAsAccessMode accessMode = XlSaveAsAccessMode.xlNoChange, object conflictResolution = null, object addToMru = null, object textCodepage = null, object textVisualLayout = null, object local = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示是否定时自动保存所有格式的更改文件。
    /// </summary>
    bool EnableAutoRecover { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否可以从工作簿中删除个人信息。默认值为False。
    /// </summary>
    bool RemovePersonalInformation { get; set; }

    /// <summary>
    /// 获取指示对象名称的字符串，包括其在磁盘上的路径。
    /// </summary>
    string FullNameURLEncoded { get; }

    /// <summary>
    /// 将工作簿从本地计算机返回到服务器，并将本地工作簿设置为只读，以便无法在本地编辑。调用此方法也会关闭工作簿。
    /// </summary>
    /// <param name="saveChanges">可选。True保存更改并签入文档。False在不保存修订的情况下将文档返回到已签入状态。</param>
    /// <param name="comments">可选。允许用户为要签入的工作簿修订输入签入注释（仅当SaveChanges等于True时适用）。</param>
    /// <param name="makePublic">可选。True允许用户在签入后发布工作簿。这将提交工作簿进行审批过程，最终可能将工作簿版本发布给具有只读权限的用户（仅当SaveChanges等于True时适用）。</param>
    void CheckIn(object saveChanges = null, object comments = null, object makePublic = null);

    /// <summary>
    /// 确定Excel是否可以将指定工作簿签入到服务器。
    /// </summary>
    /// <returns>如果可以签入则为True，否则为False。</returns>
    bool CanCheckIn();

    /// <summary>
    /// 将工作簿以电子邮件形式发送给指定收件人进行审阅。
    /// </summary>
    /// <param name="recipients">可选。列出要向其发送消息的人员的字符串。可以是电子邮件电话簿中未解析的名称和别名或完整的电子邮件地址。多个收件人之间用分号分隔。</param>
    /// <param name="subject">可选。消息主题的字符串。如果留空，主题将是：请审阅"文件名"。</param>
    /// <param name="showMessage">可选。指示执行方法时是否应显示消息的布尔值。默认值为True。</param>
    /// <param name="includeAttachment">可选。指示消息是否应包含附件或指向服务器位置的链接的布尔值。默认值为True。</param>
    void SendForReview(object recipients = null, object subject = null, object showMessage = null, object includeAttachment = null);

    /// <summary>
    /// 向已发送出去进行审阅的工作簿的作者发送电子邮件，通知他们审阅者已完成对工作簿的审阅。
    /// </summary>
    /// <param name="showMessage">可选。False不显示消息。True显示消息。</param>
    void ReplyWithChanges(object showMessage = null);

    /// <summary>
    /// 终止使用SendForReview方法发送出去进行审阅的文件审阅。
    /// </summary>
    void EndReview();

    /// <summary>
    /// 获取或设置打开指定工作簿必须提供的密码。
    /// </summary>
    string Password { get; set; }

    /// <summary>
    /// 获取或设置工作簿的写密码字符串。
    /// </summary>
    string WritePassword { get; set; }

    /// <summary>
    /// 获取指定Excel在对工作簿密码进行加密时使用的算法加密提供程序的名称。
    /// </summary>
    string PasswordEncryptionProvider { get; }

    /// <summary>
    /// 获取指示Excel用于加密工作簿密码的算法的字符串。
    /// </summary>
    string PasswordEncryptionAlgorithm { get; }

    /// <summary>
    /// 获取指示Excel加密工作簿密码时使用的算法密钥长度的整数。
    /// </summary>
    int PasswordEncryptionKeyLength { get; }

    /// <summary>
    /// 设置使用密码加密工作簿的选项。
    /// </summary>
    /// <param name="passwordEncryptionProvider">可选。加密提供程序的区分大小写的字符串。</param>
    /// <param name="passwordEncryptionAlgorithm">可选。算法短名称的区分大小写的字符串（例如"RC4"）。</param>
    /// <param name="passwordEncryptionKeyLength">可选。加密密钥长度，为8的倍数（40或更大）。</param>
    /// <param name="passwordEncryptionFileProperties">可选。True（默认）表示加密文件属性。</param>
    void SetPasswordEncryptionOptions(object passwordEncryptionProvider = null, object passwordEncryptionAlgorithm = null, object passwordEncryptionKeyLength = null, object passwordEncryptionFileProperties = null);

    /// <summary>
    /// 获取一个布尔值，表示Excel是否为指定的受密码保护的工作簿加密文件属性。
    /// </summary>
    bool PasswordEncryptionFileProperties { get; }

    /// <summary>
    /// 获取一个布尔值，表示工作簿是否以建议只读方式保存。
    /// </summary>
    bool ReadOnlyRecommended { get; set; }

    /// <summary>
    /// 保护工作簿，使其无法修改。
    /// </summary>
    /// <param name="password">可选。指定工作表或工作簿的区分大小写密码的字符串。如果省略此参数，则无需密码即可取消保护工作表或工作簿。否则，必须指定密码才能取消保护工作表或工作簿。</param>
    /// <param name="structure">可选。True表示保护工作簿的结构（工作表的相对位置）。默认值为False。</param>
    /// <param name="windows">可选。True表示保护工作簿窗口。如果省略此参数，则窗口不受保护。</param>
    void Protect(object password = null, object structure = null, object windows = null);

    /// <summary>
    /// 导致发生前台智能标记检查，自动注释以前未注释的数据。
    /// </summary>
    void RecheckSmartTags();

    /// <summary>
    /// 发送工作表作为传真给指定的收件人。
    /// </summary>
    /// <param name="recipients">可选。表示传真号码和电子邮件地址的字符串。多个收件人之间用分号分隔。</param>
    /// <param name="subject">可选。表示传真文档主题行的字符串。</param>
    /// <param name="showMessage">可选。True在发送前显示传真消息。False不显示传真消息直接发送。</param>
    void SendFaxOverInternet(object recipients = null, object subject = null, object showMessage = null);

    /// <summary>
    /// 导入XML数据文件到当前工作簿。
    /// </summary>
    /// <param name="url">必需。XML数据文件的URL或UNC路径。</param>
    /// <param name="importMap">必需。导入文件时应用的架构映射。</param>
    /// <param name="overwrite">可选。如果未为Destination参数指定值，则此参数指定是否覆盖映射到ImportMap参数中指定的架构映射的数据。设置为True表示覆盖数据，False表示将新数据附加到现有数据。默认值为True。</param>
    /// <param name="destination">可选。数据将导入到指定范围的新XML列表中。</param>
    /// <returns>XML导入结果。</returns>
    XlXmlImportResult XmlImport(string url, out XmlMap importMap, object overwrite = null, object destination = null);

    /// <summary>
    /// 导出已映射到指定XML架构映射的数据到XML数据文件。
    /// </summary>
    /// <param name="filename">必需。指示要保存的文件名的字符串。可以包含完整路径；如果不包含，Microsoft Excel将文件保存在当前文件夹中。</param>
    /// <param name="map">必需。应用于数据的架构映射。</param>
    void SaveAsXMLData(string filename, XmlMap map);

    /// <summary>
    /// 打开或关闭窗体设计模式。
    /// </summary>
    void ToggleFormsDesign();

    /// <summary>
    /// 删除工作簿中指定类型的所有信息。
    /// </summary>
    /// <param name="removeDocInfoType">指示要删除信息类型的XlRemoveDocInfoType值之一。</param>
    void RemoveDocumentInformation(XlRemoveDocInfoType removeDocInfoType);

    /// <summary>
    /// 将工作簿从本地计算机保存到服务器，并将本地工作簿设置为只读，使其无法在本地编辑。
    /// </summary>
    /// <param name="saveChanges">True将工作簿保存到服务器位置。默认为True。</param>
    /// <param name="comments">要签入的工作簿修订的注释（仅当SaveChanges设置为True时适用）。</param>
    /// <param name="makePublic">True允许用户在签入后发布工作簿。</param>
    /// <param name="versionType">指定工作簿的版本控制信息。</param>
    void CheckInWithVersion(object saveChanges = null, object comments = null, object makePublic = null, object versionType = null);

    /// <summary>
    /// 锁定服务器上的工作簿以防止修改。
    /// </summary>
    void LockServerFile();

    /// <summary>
    /// 以PDF或XPS格式发布工作簿。
    /// </summary>
    /// <param name="type">可以是xlTypePDF或xlTypeXPS。</param>
    /// <param name="filename">指示要保存的文件名的字符串。可以包含完整路径或短名称。Excel2007将文件保存在当前文件夹中。</param>
    /// <param name="quality">可以设置为xlQualityStandard或xlQualityMinimum。</param>
    /// <param name="includeDocProperties">设置为True表示包含文档属性，False表示省略。</param>
    /// <param name="ignorePrintAreas">如果设置为True，则发布时忽略任何设置的打印区域。如果设置为False，则发布时使用设置的打印区域。</param>
    /// <param name="from">开始发布的页码。如果省略此参数，则从开头开始发布。</param>
    /// <param name="to">要发布的最后一页的页码。如果省略此参数，则发布到最后一页。</param>
    /// <param name="openAfterPublish">如果设置为True，发布后在查看器中显示文件。如果设置为False，则发布文件但不显示。</param>
    /// <param name="fixedFormatExtClassPtr">FixedFormatExt类的指针。</param>
    void ExportAsFixedFormat(XlFixedFormatType type, object filename = null, object quality = null, object includeDocProperties = null, object ignorePrintAreas = null, object from = null, object to = null, object openAfterPublish = null, object fixedFormatExtClassPtr = null);

    /// <summary>
    /// 获取或设置Excel在对文档进行加密时使用的算法加密提供程序的名称。
    /// </summary>
    string EncryptionProvider { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示如果工作簿包含Excel早期版本不支持的功能，是否应提示用户转换工作簿。
    /// </summary>
    bool DoNotPromptForConvert { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否强制对工作簿进行完整计算。
    /// </summary>
    bool ForceFullCalculation { get; set; }

    /// <summary>
    /// 获取与工作簿关联的切片器缓存。
    /// </summary>
    SlicerCaches SlicerCaches { get; }

    /// <summary>
    /// 获取活动工作簿或指定工作簿中的活动切片器。
    /// </summary>
    Slicer ActiveSlicer { get; }

    /// <summary>
    /// 获取或设置用作切片器默认样式的样式。
    /// </summary>
    object DefaultSlicerStyle { get; set; }

    /// <summary>
    /// 获取或设置某些工作表函数是否使用最新的精度算法来计算结果。
    /// </summary>
    int AccuracyVersion { get; set; }

    /// <summary>
    /// 获取数据模型。
    /// </summary>
    Model Model { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否跟踪图表数据点。
    /// </summary>
    bool ChartDataPointTrack { get; set; }

    /// <summary>
    /// 获取或设置用作时间线默认样式的样式。
    /// </summary>
    object DefaultTimelineStyle { get; set; }

    /// <summary>
    /// 获取或设置搜索时是否区分大小写。
    /// </summary>
    bool CaseSensitive { get; }

    /// <summary>
    /// 获取或设置搜索时是否使用全单元格条件。
    /// </summary>
    bool UseWholeCellCriteria { get; }

    /// <summary>
    /// 获取或设置搜索时是否使用通配符。
    /// </summary>
    bool UseWildcards { get; }

    /// <summary>
    /// 获取数据透视表集合。
    /// </summary>
    object PivotTables { get; }

    #region 事件
    /// <summary>
    /// 当工作表数据透视表更改时发生
    /// </summary>
    event WorkBookSheetPivotTableChangeEventHandler WorkBookSheetPivotTableChange;

    /// <summary>
    /// 当窗口大小调整时发生
    /// </summary>
    event WindowResizeEventHandler WindowResize;

    /// <summary>
    /// 当窗口停用时发生
    /// </summary>
    event WindowDeactivateEventHandler WindowDeActivate;

    /// <summary>
    /// 当窗口激活时发生
    /// </summary>
    event WindowActivateEventHandler WindowActivate;

    /// <summary>
    /// 当工作表计算完成时发生
    /// </summary>
    event SheetCalculateEventHandler Calculate;

    /// <summary>
    /// 当工作表即将删除时发生
    /// </summary>
    event SheetBeforeDeleteEventHandler SheetBeforeDelete;

    /// <summary>
    /// 当工作表停用时发生
    /// </summary>
    event SheetDeactivateEventHandler SheetDeactivate;

    /// <summary>
    /// 当行集完成之前发生
    /// </summary>
    event WorkBookBeforeRowsetCompleteEventHandler OnBeforeRowsetComplete;

    /// <summary>
    /// 当数据透视表打开连接时发生
    /// </summary>
    event WorkBookPivotTableOpenConnectionEventHandler PivotTableOpenConnection;

    /// <summary>
    /// 当数据透视表关闭连接时发生
    /// </summary>
    event WorkBookPivotTableCloseConnectionEventHandler PivotTableCloseConnection;

    /// <summary>
    /// 当工作簿打开时发生
    /// </summary>
    event WorkbookOpenEventHandler Open;

    /// <summary>
    /// 当工作簿中新建工作表时发生
    /// </summary>
    event WorkBookNewSheetEventHandler NewSheet;

    /// <summary>
    /// 当工作簿停用时发生
    /// </summary>
    event DeactivateEventHandler Deactivate;

    /// <summary>
    /// 当工作簿打印之前发生
    /// </summary>
    event WorkBookBeforePrintEventHandler BeforePrint;

    /// <summary>
    /// 当工作簿中新建图表时发生
    /// </summary>
    event WorkBookNewChartEventHandler NewChart;

    /// <summary>
    /// 当工作簿关闭之前发生
    /// </summary>
    event WorkbookBeforeCloseEventHandler BeforeClose;

    /// <summary>
    /// 当工作表选择区域更改时发生
    /// </summary>
    event SheetSelectionChangeEventHandler SheetSelectionChange;

    /// <summary>
    /// 当工作簿保存之后发生
    /// </summary>
    event WorkBookAfterSaveEventHandler AfterSave;

    /// <summary>
    /// 当工作簿激活时发生
    /// </summary>
    event WorkbookActivateEventHandler WorkbookActivate;

    /// <summary>
    /// 当工作表内容更改时发生
    /// </summary>
    event WorkBookSheetChangeEventHandler SheetChange;

    /// <summary>
    /// 当工作簿同步事件发生时触发
    /// </summary>
    event WorkBookSyncEventHandler Sync;
    #endregion
}
