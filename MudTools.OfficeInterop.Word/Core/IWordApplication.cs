//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Vbe;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word应用程序接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord", NoneConstructor = true, NoneDisposed = true)]
public partial interface IWordApplication : IOfficeObject<IWordApplication, MsWord.Application>, IOfficeApplication
{
    /// <summary>
    /// 获取表示 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取一个 32 位整数，指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取表示指定对象的父对象的对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取表示所有打开文档的 Documents 集合。
    /// </summary>
    IWordDocuments? Documents { get; }

    /// <summary>
    /// 获取表示所有文档窗口的 Windows 集合。
    /// </summary>
    IWordWindows? Windows { get; }

    /// <summary>
    /// 获取表示活动文档的 Document 对象。
    /// </summary>
    IWordDocument? ActiveDocument { get; }

    /// <summary>
    /// 获取表示活动窗口的 Window 对象。
    /// </summary>
    IWordWindow? ActiveWindow { get; }

    /// <summary>
    /// 获取表示选定范围或插入点的 Selection 对象。
    /// </summary>
    IWordSelection? Selection { get; }

    /// <summary>
    /// 获取包含 Microsoft Word 6.0 和 Word for Windows 95 中所有可用 WordBasic 语句和函数的自动化对象。
    /// </summary>
    object WordBasic { get; }

    /// <summary>
    /// 获取表示最近访问文件的 RecentFiles 集合。
    /// </summary>
    IWordRecentFiles? RecentFiles { get; }

    /// <summary>
    /// 获取表示 Normal 模板的 Template 对象。
    /// </summary>
    IWordTemplate? NormalTemplate { get; }

    /// <summary>
    /// 获取可用于返回系统相关信息并执行系统相关任务的 System 对象。
    /// </summary>
    IWordSystem? System { get; }

    /// <summary>
    /// 获取包含当前自动更正选项、条目和例外的 AutoCorrect 对象。
    /// </summary>
    IWordAutoCorrect? AutoCorrect { get; }

    /// <summary>
    /// 获取包含所有可用字体名称的 FontNames 对象。
    /// </summary>
    IWordFontNames? FontNames { get; }

    /// <summary>
    /// 获取包含所有可用横向字体名称的 FontNames 对象。
    /// </summary>
    IWordFontNames? LandscapeFontNames { get; }

    /// <summary>
    /// 获取包含所有可用纵向字体名称的 FontNames 对象。
    /// </summary>
    IWordFontNames? PortraitFontNames { get; }

    /// <summary>
    /// 获取表示"语言"对话框中所列校对语言的 Languages 集合。
    /// </summary>
    IWordLanguages? Languages { get; }

    /// <summary>
    /// 获取表示垂直滚动条上"选择浏览对象"工具的 Browser 对象。
    /// </summary>
    IWordBrowser? Browser { get; }

    /// <summary>
    /// 获取表示 Microsoft Word 可用的所有文件转换器的 FileConverters 集合。
    /// </summary>
    IWordFileConverters? FileConverters { get; }

    /// <summary>
    /// 获取表示邮件标签的 MailingLabel 对象。
    /// </summary>
    IWordMailingLabel? MailingLabel { get; }

    /// <summary>
    /// 获取表示 Microsoft Word 中所有内置对话框的 Dialogs 集合。
    /// </summary>
    IWordDialogs? Dialogs { get; }

    /// <summary>
    /// 获取表示所有可用题注标签的 CaptionLabels 集合。
    /// </summary>
    IWordCaptionLabels? CaptionLabels { get; }

    /// <summary>
    /// 获取表示插入表格和图片等项时自动添加的题注的 AutoCaptions 集合。
    /// </summary>
    IWordAutoCaptions? AutoCaptions { get; }

    /// <summary>
    /// 获取表示所有可用加载项（无论是否当前加载）的 AddIns 集合。
    /// </summary>
    IWordAddIns? AddIns { get; }

    /// <summary>
    /// 获取或设置屏幕更新是否打开。返回 True 表示屏幕更新已打开，False 表示未打开。
    /// </summary>
    bool ScreenUpdating { get; set; }

    /// <summary>
    /// 获取或设置打印预览是否为当前视图。返回 True 表示打印预览是当前视图，False 表示不是。
    /// </summary>
    bool PrintPreview { get; set; }

    /// <summary>
    /// 获取表示所有正在运行的应用程序的 Tasks 集合。
    /// </summary>
    IWordTasks? Tasks { get; }

    /// <summary>
    /// 获取或设置一个值，指示状态栏是否显示。
    /// </summary>
    bool DisplayStatusBar { get; set; }

    /// <summary>
    /// 获取 Microsoft Word 是否处于特殊模式（例如复制文本模式或移动文本模式）。
    /// </summary>
    bool SpecialMode { get; }

    /// <summary>
    /// 获取可设置为 Microsoft Word 文档窗口的最大宽度（以磅为单位）。
    /// </summary>
    int UsableWidth { get; }

    /// <summary>
    /// 获取可设置为 Microsoft Word 文档窗口的最大高度（以磅为单位）。
    /// </summary>
    int UsableHeight { get; }

    /// <summary>
    /// 获取数学协处理器是否已安装并可供 Microsoft Word 使用。
    /// 返回 True 表示数学协处理器已安装并可用，False 表示不可用。
    /// </summary>
    bool MathCoprocessorAvailable { get; }

    /// <summary>
    /// 获取系统是否可用鼠标。返回 True 表示系统有可用鼠标，False 表示没有。
    /// </summary>
    bool MouseAvailable { get; }

    /// <summary>
    /// 获取 CAPS LOCK 键是否打开。返回 True 表示 CAPS LOCK 键已打开，False 表示未打开。
    /// </summary>
    bool CapsLock { get; }

    /// <summary>
    /// 获取 NUM LOCK 键的状态。返回 True 表示数字键盘上的键插入数字，False 表示这些键移动插入点。
    /// </summary>
    bool NumLock { get; }

    /// <summary>
    /// 获取或设置用户的姓名，用于信封和"作者"文档属性。
    /// </summary>
    string UserName { get; set; }

    /// <summary>
    /// 获取或设置用户的缩写，Microsoft Word 使用它来构造批注标记。
    /// </summary>
    string UserInitials { get; set; }

    /// <summary>
    /// 获取或设置用户的邮件地址。
    /// </summary>
    string UserAddress { get; set; }

    /// <summary>
    /// 获取表示存储运行过程的模块的模板或文档的 Template 或 Document 对象。
    /// </summary>
    object MacroContainer { get; }

    /// <summary>
    /// 获取或设置是否在"文件"菜单中显示最近使用的文件名。
    /// </summary>
    bool DisplayRecentFiles { get; set; }

    /// <summary>
    /// 获取包含指定单词或短语的同义词、反义词或相关单词和表达式的同义词库信息的 SynonymInfo 对象。
    /// </summary>
    [MethodIndex]
    IWordSynonymInfo? SynonymInfo(string word, [ComNamespace("MsCore")] MsoLanguageID languageID);

    /// <summary>
    /// 获取表示 Visual Basic 编辑器的 VBE 对象。
    /// </summary>
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取或设置"另存为"对话框（"文件"菜单）中"保存类型"框中显示的默认格式。
    /// </summary>
    string DefaultSaveFormat { get; set; }

    /// <summary>
    /// 获取表示三个列表模板库（项目符号、编号和多级编号）的 ListGalleries 集合。
    /// </summary>
    IWordListGalleries? ListGalleries { get; }

    /// <summary>
    /// 获取或设置活动打印机的名称。
    /// </summary>
    string ActivePrinter { get; set; }

    /// <summary>
    /// 获取表示所有可用模板（全局模板以及附加到打开文档的模板）的 Templates 集合。
    /// </summary>
    IWordTemplates? Templates { get; }

    /// <summary>
    /// 获取或设置存储菜单栏、工具栏和键绑定更改的模板或文档的 Template 或 Document 对象。
    /// </summary>
    object CustomizationContext { get; set; }

    /// <summary>
    /// 获取表示自定义键分配的 KeyBindings 集合，包括键代码、键类别和命令。
    /// </summary>
    IWordKeyBindings? KeyBindings { get; }

    /// <summary>
    /// 获取表示指定组合键的 KeyBinding 对象。
    /// </summary>
    [MethodIndex]
    IWordKeyBinding? FindKey(int keyCode, int? keyCode2);

    /// <summary>
    /// 获取或设置指定文档或应用程序窗口的标题文本。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置是否在至少一个文档窗口中显示滚动条。
    /// </summary>
    bool DisplayScrollBars { get; set; }

    /// <summary>
    /// 获取或设置启动文件夹的完整路径，不包括最终分隔符。
    /// </summary>
    string StartupPath { get; set; }

    /// <summary>
    /// 获取排队等待在后台保存的文件数量。
    /// </summary>
    int BackgroundSavingStatus { get; }

    /// <summary>
    /// 获取后台打印队列中的打印作业数量。
    /// </summary>
    int BackgroundPrintingStatus { get; }

    /// <summary>
    /// 获取或设置指定文档窗口或任务窗口的状态。
    /// </summary>
    WdWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Word 是否显示建议文本以完成输入时的单词、日期或短语。
    /// </summary>
    bool DisplayAutoCompleteTips { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Word 中应用程序设置的 Options 对象。
    /// </summary>
    IWordOptions? Options { get; }

    /// <summary>
    /// 获取或设置宏运行时处理某些警报和消息的方式。
    /// </summary>
    WdAlertLevel DisplayAlerts { get; set; }

    /// <summary>
    /// 获取表示活动自定义词典集合的 Dictionaries 对象。
    /// </summary>
    IWordDictionaries? CustomDictionaries { get; }

    /// <summary>
    /// 获取用于分隔文件夹名称的字符。
    /// </summary>
    string PathSeparator { get; }

    /// <summary>
    /// 在状态栏中显示指定的文本。
    /// </summary>
    string StatusBar { set; }

    /// <summary>
    /// 获取是否安装了 MAPI。返回 True 表示已安装 MAPI，False 表示未安装。
    /// </summary>
    bool MAPIAvailable { get; }

    /// <summary>
    /// 获取或设置是否将批注、脚注、尾注和超链接显示为提示。标记为有批注的文本会突出显示。
    /// </summary>
    bool DisplayScreenTips { get; set; }

    /// <summary>
    /// 获取或设置 Word 处理 CTRL+BREAK 用户中断的方式。
    /// </summary>
    WdEnableCancelKey EnableCancelKey { get; set; }

    /// <summary>
    /// 获取文档或应用程序是否由用户创建或打开。
    /// </summary>
    bool UserControl { get; }

    /// <summary>
    /// 获取可用于使用绝对或相对路径搜索文件的 FileSearch 对象。
    /// </summary>
    IOfficeFileSearch? FileSearch { get; }

    /// <summary>
    /// 获取主机上安装的邮件系统（或多个系统）。
    /// </summary>
    WdMailSystem MailSystem { get; }

    /// <summary>
    /// 获取或设置将文本转换为表格时用于将文本分隔成单元格的单个字符。
    /// </summary>
    string DefaultTableSeparator { get; set; }

    /// <summary>
    /// 获取或设置 Visual Basic 编辑器窗口是否可见。返回 True 表示 Visual Basic 编辑器窗口可见，False 表示不可见。
    /// </summary>
    bool ShowVisualBasicEditor { get; set; }

    /// <summary>
    /// 设置此属性为 "text/html" 以允许在 Microsoft Word 中打开超链接的 HTML 文件（而不是默认的 Internet 浏览器）。
    /// </summary>
    string BrowseExtraFileTypes { get; set; }

    /// <summary>
    /// 获取引用对象的指定变量是否有效。
    /// </summary>
    [MethodIndex]
    bool? IsObjectValid(object? variable);

    /// <summary>
    /// 获取表示所有活动自定义转换词典的 HangulHanjaConversionDictionaries 集合。
    /// </summary>
    IWordHangulHanjaConversionDictionaries? HangulHanjaDictionaries { get; }

    /// <summary>
    /// 获取表示活动电子邮件的 MailMessage 对象。
    /// </summary>
    IWordMailMessage? MailMessage { get; }

    /// <summary>
    /// 获取插入点是否在电子邮件标题字段中。
    /// </summary>
    bool FocusInMailHeader { get; }

    /// <summary>
    /// 获取表示 Microsoft Word 用户界面所选语言的 MsoLanguageID 常量。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLanguageID Language { get; }

    /// <summary>
    /// 获取或设置 Microsoft Word 是否在键入时自动检测您使用的语言。
    /// 返回 True 表示 Microsoft Word 自动检测您使用的语言，False 表示不自动检测。
    /// </summary>
    bool CheckLanguage { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Word 如何处理调用尚未安装功能的方法和属性。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoFeatureInstall FeatureInstall { get; set; }

    /// <summary>
    /// 获取表示电子邮件撰写全局首选项的 EmailOptions 对象。
    /// </summary>
    IWordEmailOptions? EmailOptions { get; }

    /// <summary>
    /// 获取表示当前加载到 Microsoft Word 中的所有组件对象模型 (COM) 加载项的 COMAddIns 集合。
    /// </summary>
    IOfficeCOMAddIns? COMAddIns { get; }

    /// <summary>
    /// 获取表示联机帮助搜索引擎使用的文件的 AnswerWizard 对象。
    /// </summary>
    //IOfficeAnswerWizard? AnswerWizard { get; } 

    /// <summary>
    /// 获取包含存储在 Microsoft Office Word 中的书目参考源的 Bibliography 对象。
    /// </summary>
    IWordBibliography? Bibliography { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Office Word 是否在"样式"对话框中显示样式的格式预览。
    /// </summary>
    bool ShowStylePreviews { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Office Word 是否允许链接样式。
    /// </summary>
    bool RestrictLinkedStyles { get; set; }

    /// <summary>
    /// 获取用于方程式的自动更正条目。
    /// </summary>
    IWordOMathAutoCorrect? OMathAutoCorrect { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示文档属性面板。
    /// </summary>
    bool DisplayDocumentInformationPanel { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Office 帮助查看器的 IAssistance 对象。
    /// </summary>
    IOfficeAssistance? Assistance { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Office Word 是否在阅读模式下打开电子邮件附件。
    /// </summary>
    bool OpenAttachmentsInFullScreen { get; set; }

    /// <summary>
    /// 获取表示与活动文档关联的加密会话的整数值。
    /// </summary>
    int ActiveEncryptionSession { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Office Word 是否在运行其他代码后保持插入点位置的格式属性。
    /// </summary>
    bool DontResetInsertionPointProperties { get; set; }

    /// <summary>
    /// 获取表示当前加载到应用程序中的 SmartArt 布局集合的 SmartArtLayouts 对象。
    /// </summary>
    IOfficeSmartArtLayouts? SmartArtLayouts { get; }

    /// <summary>
    /// 获取表示当前加载到应用程序中的 SmartArt 样式集合的 SmartArtQuickStyles 对象。
    /// </summary>
    IOfficeSmartArtQuickStyles? SmartArtQuickStyles { get; }

    /// <summary>
    /// 获取表示当前加载到应用程序中的颜色样式集合的 SmartArtColors 对象。
    /// </summary>
    IOfficeSmartArtColors? SmartArtColors { get; }

    /// <summary>
    /// 获取提供对撤消堆栈自定义入口点的 UndoRecord 对象。
    /// </summary>
    IWordUndoRecord? UndoRecord { get; }

    /// <summary>
    /// 获取提供在对话框中选择人员或数据功能的 PickerDialog 对象。
    /// </summary>
    IOfficePickerDialog? PickerDialog { get; }

    /// <summary>
    /// 获取表示所有受保护的视图窗口的 ProtectedViewWindows 集合。
    /// </summary>
    IWordProtectedViewWindows? ProtectedViewWindows { get; }

    /// <summary>
    /// 获取表示活动受保护的视图窗口的 ProtectedViewWindow 对象。
    /// </summary>
    IWordProtectedViewWindow? ActiveProtectedViewWindow { get; }

    /// <summary>
    /// 获取应用程序窗口是否为受保护的视图窗口。
    /// 如果应用程序窗口是受保护的视图窗口，则为 true；否则为 false。
    /// </summary>
    bool IsSandboxed { get; }

    /// <summary>
    /// 获取或设置在打开文件之前 Word 如何验证文件。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoFileValidationMode FileValidation { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否跟踪图表数据点。
    /// </summary>
    bool ChartDataPointTrack { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示动画。
    /// </summary>
    bool ShowAnimation { get; set; }

    /// <summary>
    /// 退出 Microsoft Word 并可选择保存或路由打开的文档。
    /// </summary>
    /// <param name="saveChanges">指定 Word 退出前是否保存已更改的文档。可以是任何 WdSaveOptions 常量。</param>
    /// <param name="originalFormat">指定 Word 保存非 Word 文档格式文档的方式。可以是任何 WdOriginalFormat 常量。</param>
    /// <param name="routeDocument">True 将文档路由到下一个收件人。如果文档没有附加路由单，则忽略此参数。</param>
    void Quit(WdSaveOptions? saveChanges = null, WdOriginalFormat? originalFormat = null, bool? routeDocument = null);

    /// <summary>
    /// 用视频内存缓冲区中的当前信息更新监视器上的显示。
    /// </summary>
    void ScreenRefresh();

    /// <summary>
    /// 在全局通讯簿列表中查找名称并显示包含指定名称信息的"属性"对话框。
    /// </summary>
    /// <param name="name">全局通讯簿中的名称。</param>
    void LookupNameProperties(string name);

    /// <summary>
    /// 设置字体映射选项，这些选项反映在"字体替换"对话框（"工具"菜单，"选项"对话框，"兼容性"选项卡）中。
    /// </summary>
    /// <param name="unavailableFont">计算机上不可用的字体名称，您希望将其映射到其他字体以进行显示和打印。</param>
    /// <param name="substituteFont">计算机上可用字体名称，您希望用它替换不可用字体。</param>
    void SubstituteFont(string unavailableFont, string substituteFont);

    /// <summary>
    /// 重复最近的编辑操作一次或多次。
    /// </summary>
    /// <param name="times">要重复上次命令的次数。</param>
    /// <returns>如果重复操作成功，则为 True；否则为 False。</returns>
    bool? Repeat(int? times = null);

    /// <summary>
    /// 通过指定的动态数据交换 (DDE) 通道向应用程序发送一个命令或一系列命令。
    /// </summary>
    /// <param name="channel">DDEInitiate 方法返回的通道号。</param>
    /// <param name="command">接收应用程序（DDE 服务器）识别的命令或一系列命令。如果接收应用程序无法执行指定命令，则会发生错误。</param>
    void DDEExecute(int channel, string command);

    /// <summary>
    /// 打开到另一个应用程序的动态数据交换 (DDE) 通道，并返回通道号。
    /// </summary>
    /// <param name="app">应用程序名称。</param>
    /// <param name="topic">DDE 主题名称，例如打开文档的名称，由您要打开通道的应用程序识别。</param>
    /// <returns>DDE 通道号。</returns>
    int? DDEInitiate(string app, string topic);

    /// <summary>
    /// 使用打开的动态数据交换 (DDE) 通道向应用程序发送数据。
    /// </summary>
    /// <param name="channel">DDEInitiate 方法返回的通道号。</param>
    /// <param name="item">DDE 主题中的项，指定数据将发送到该项。</param>
    /// <param name="data">要发送到接收应用程序（DDE 服务器）的数据。</param>
    void DDEPoke(int channel, string item, string data);

    /// <summary>
    /// 使用打开的动态数据交换 (DDE) 通道向接收应用程序请求信息，并将信息作为字符串返回。
    /// </summary>
    /// <param name="channel">DDEInitiate 方法返回的通道号。</param>
    /// <param name="item">要请求的项。</param>
    /// <returns>请求的信息字符串。</returns>
    string? DDERequest(int channel, string item);

    /// <summary>
    /// 关闭到另一个应用程序的指定动态数据交换 (DDE) 通道。
    /// </summary>
    /// <param name="channel">DDEInitiate 方法返回的通道号。</param>
    void DDETerminate(int channel);

    /// <summary>
    /// 关闭由 Microsoft Word 打开的所有动态数据交换 (DDE) 通道。
    /// </summary>
    void DDETerminateAll();

    /// <summary>
    /// 返回指定组合键的唯一编号。
    /// </summary>
    /// <param name="arg1">使用 WdKey 常量之一指定的键。</param>
    /// <param name="arg2">使用 WdKey 常量之一指定的键。</param>
    /// <param name="arg3">使用 WdKey 常量之一指定的键。</param>
    /// <param name="arg4">使用 WdKey 常量之一指定的键。</param>
    /// <returns>组合键的唯一编号。</returns>
    int? BuildKeyCode(WdKey arg1, WdKey? arg2 = null, WdKey? arg3 = null, WdKey? arg4 = null);

    /// <summary>
    /// 返回指定键的键组合字符串（例如，CTRL+SHIFT+A）。
    /// </summary>
    /// <param name="keyCode">使用 WdKey 常量之一指定的键。</param>
    /// <param name="keyCode2">使用 WdKey 常量之一指定的第二个键。</param>
    /// <returns>键组合字符串。</returns>
    string? KeyString(int keyCode, WdKey? keyCode2 = null);

    /// <summary>
    /// 将指定的自动图文集条目、工具栏、样式或宏项目项从源文档或模板复制到目标文档或模板。
    /// </summary>
    /// <param name="source">包含要复制的项的文档或模板文件名。</param>
    /// <param name="destination">要将项复制到的文档或模板文件名。</param>
    /// <param name="name">要复制的自动图文集条目、工具栏、样式或宏的名称。</param>
    /// <param name="objectType">要复制的项的类型。</param>
    void OrganizerCopy(string source, string destination, string name, WdOrganizerObject objectType);

    /// <summary>
    /// 从文档或模板中删除指定的样式、自动图文集条目、工具栏或宏项目项。
    /// </summary>
    /// <param name="source">包含要删除项的文档或模板文件名。</param>
    /// <param name="name">要删除的样式、自动图文集条目、工具栏或宏的名称。</param>
    /// <param name="objectType">要删除的项的类型。</param>
    void OrganizerDelete(string source, string name, WdOrganizerObject objectType);

    /// <summary>
    /// 重命名文档或模板中的指定样式、自动图文集条目、工具栏或宏项目项。
    /// </summary>
    /// <param name="source">包含要重命名项的文档或模板文件名。</param>
    /// <param name="name">要重命名的样式、自动图文集条目、工具栏或宏的名称。</param>
    /// <param name="newName">项的新名称。</param>
    /// <param name="objectType">要重命名的项的类型。</param>
    void OrganizerRename(string source, string name, string newName, WdOrganizerObject objectType);

    /// <summary>
    /// 向通讯簿添加条目。
    /// </summary>
    /// <param name="tagID">新地址条目的标记 ID 值数组。</param>
    /// <param name="value">新地址条目的值数组。每个元素对应于 TagID 数组中的一个元素。</param>
    void AddAddress(Array tagID, Array value);

    /// <summary>
    /// 从默认通讯簿返回地址。
    /// </summary>
    /// <param name="name">收件人姓名，如通讯簿中"搜索姓名"对话框所示。</param>
    /// <param name="addressProperties">
    /// 如果 UseAutoText 为 True，此参数表示定义地址簿属性序列的自动图文集条目的名称。
    /// 如果 UseAutoText 为 False 或省略，此参数定义自定义布局。
    /// </param>
    /// <param name="useAutoText">True 表示 AddressProperties 指定定义地址簿属性序列的自动图文集条目的名称；False 表示它指定自定义布局。</param>
    /// <param name="displaySelectDialog">指定是否显示"选择姓名"对话框。</param>
    /// <param name="selectDialog">指定"选择姓名"对话框的显示方式（即模式）。</param>
    /// <param name="checkNamesDialog">如果 Name 参数不够具体，则为 True 显示"检查姓名"对话框。</param>
    /// <param name="recentAddressesChoice">True 使用最近使用的返回地址列表。</param>
    /// <param name="updateRecentAddresses">True 将地址添加到最近使用的地址列表；False 不添加地址。</param>
    /// <returns>通讯簿地址字符串。</returns>
    string? GetAddress(string? name = null, object? addressProperties = null, bool? useAutoText = null,
                      bool? displaySelectDialog = null, object? selectDialog = null, bool? checkNamesDialog = null,
                      bool? recentAddressesChoice = null, bool? updateRecentAddresses = null);

    /// <summary>
    /// 检查字符串的语法错误。
    /// </summary>
    /// <param name="text">要检查语法错误的字符串。</param>
    /// <returns>如果字符串没有语法错误，则为 True；否则为 False。</returns>
    bool? CheckGrammar(string text);

    /// <summary>
    /// 检查字符串的拼写错误。
    /// </summary>
    /// <param name="word">要检查拼写的文本。</param>
    /// <param name="customDictionary">返回 Dictionary 对象的表达式或自定义词典的文件名。</param>
    /// <param name="ignoreUppercase">True 忽略大写。如果省略此参数，则使用 Options.IgnoreUppercase 属性的当前值。</param>
    /// <param name="mainDictionary">返回 Dictionary 对象的表达式或主词典的文件名。</param>
    /// <param name="customDictionary2">返回 Dictionary 对象的表达式或附加自定义词典的文件名。最多可指定九个附加词典。</param>
    /// <param name="customDictionary3">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary4">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary5">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary6">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary7">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary8">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary9">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary10">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <returns>如果单词拼写正确，则为 True；否则为 False。</returns>
    bool? CheckSpelling(string word, IWordDictionary? customDictionary = null, IWordDictionary? ignoreUppercase = null,
                        IWordDictionary? mainDictionary = null, IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null,
                        IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null, IWordDictionary? customDictionary6 = null,
                        IWordDictionary? customDictionary7 = null, IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null,
                        IWordDictionary? customDictionary10 = null);

    /// <summary>
    /// 清除之前在进行拼写检查时忽略的单词列表。
    /// </summary>
    void ResetIgnoreAll();

    /// <summary>
    /// 返回表示建议作为给定单词拼写替换的单词的 SpellingSuggestions 集合。
    /// </summary>
    /// <param name="word">要检查拼写的单词。</param>
    /// <param name="customDictionary">返回 Dictionary 对象的表达式或自定义词典的文件名。</param>
    /// <param name="ignoreUppercase">True 忽略所有大写字母的单词。</param>
    /// <param name="mainDictionary">返回 Dictionary 对象的表达式或主词典的文件名。</param>
    /// <param name="suggestionMode">指定 Word 提供拼写建议的方式。</param>
    /// <param name="customDictionary2">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary3">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary4">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary5">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary6">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary7">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary8">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary9">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <param name="customDictionary10">返回 Dictionary 对象的表达式或附加自定义词典的文件名。</param>
    /// <returns>拼写建议集合。</returns>
    IWordSpellingSuggestions? GetSpellingSuggestions(string word, IWordDictionary? customDictionary = null, IWordDictionary? ignoreUppercase = null,
                                    IWordDictionary? mainDictionary = null, IWordDictionary? suggestionMode = null, IWordDictionary? customDictionary2 = null,
                                    IWordDictionary? customDictionary3 = null, IWordDictionary? customDictionary4 = null, IWordDictionary? customDictionary5 = null,
                                    IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null, IWordDictionary? customDictionary8 = null,
                                    IWordDictionary? customDictionary9 = null, IWordDictionary? customDictionary10 = null);

    /// <summary>
    /// 在活动文档中最近编辑发生的三个位置之间移动插入点。
    /// </summary>
    void GoBack();

    /// <summary>
    /// 显示联机帮助信息。
    /// </summary>
    /// <param name="helpType">联机帮助主题或窗口。可以是任何 WdHelpType 常量。</param>
    void Help(object helpType);

    /// <summary>
    /// 当 Office 助手建议更改时，执行自动套用格式操作。
    /// </summary>
    void AutomaticChange();

    /// <summary>
    /// 当有更多信息可用时，显示 Office 助手或帮助窗口。
    /// </summary>
    void ShowMe();

    /// <summary>
    /// 将指针从箭头更改为问号，指示您将获得有关下一个命令或屏幕元素的上下文相关帮助信息。
    /// </summary>
    void HelpTool();

    /// <summary>
    /// 打开与指定窗口具有相同文档的新窗口。
    /// </summary>
    /// <returns>新窗口对象。</returns>
    IWordWindow? NewWindow();

    /// <summary>
    /// 创建一个新文档，然后插入 Microsoft Word 命令表及其关联的快捷键和菜单分配。
    /// </summary>
    /// <param name="listAllCommands">True 包括所有 Word 命令及其分配（无论是自定义的还是内置的）；False 仅包括具有自定义分配的命令。</param>
    void ListCommands(bool listAllCommands);

    /// <summary>
    /// 显示剪贴板任务窗格。
    /// </summary>
    void ShowClipboard();

    /// <summary>
    /// 启动在指定时间运行宏的后台计时器。
    /// </summary>
    /// <param name="when">运行宏的时间。可以是表示时间的字符串，也可以是函数返回的序列号。</param>
    /// <param name="name">要运行的宏的名称。使用完整的宏路径以确保运行正确的宏。</param>
    /// <param name="tolerance">在取消未在 When 指定时间运行的宏之前可以经过的最大时间（秒）。</param>
    void OnTime(object when, string name, int? tolerance = null);

    /// <summary>
    /// 从字符串中删除非打印字符（字符代码 1-29）和特殊 Microsoft Word 字符，或将它们更改为空格（字符代码 32）。
    /// </summary>
    /// <param name="text">源字符串。</param>
    /// <returns>清理后的字符串。</returns>
    string CleanString(string text);

    /// <summary>
    /// 设置 Microsoft Word 搜索文档的文件夹。
    /// </summary>
    /// <param name="path">Word 搜索文档的文件夹路径。</param>
    void ChangeFileOpenDirectory(string path);

    /// <summary>
    /// 在活动文档中向前移动插入点至最近编辑发生的三个位置之间。
    /// </summary>
    void GoForward();

    /// <summary>
    /// 定位任务窗口或活动文档窗口。
    /// </summary>
    /// <param name="left">指定窗口的水平屏幕位置。</param>
    /// <param name="top">指定窗口的垂直屏幕位置。</param>
    void Move(int left, int top);

    /// <summary>
    /// 调整 Microsoft Word 应用程序窗口或指定任务窗口的大小。
    /// </summary>
    /// <param name="width">窗口的宽度（以磅为单位）。</param>
    /// <param name="height">窗口的高度（以磅为单位）。</param>
    void Resize(int width, int height);

    /// <summary>
    /// 将测量值从英寸转换为磅（1 英寸 = 72 磅）。
    /// </summary>
    /// <param name="inches">要转换为磅的英寸值。</param>
    /// <returns>转换后的磅值。</returns>
    float? InchesToPoints(float inches);

    /// <summary>
    /// 将测量值从厘米转换为磅（1 厘米 = 28.35 磅）。
    /// </summary>
    /// <param name="centimeters">要转换为磅的厘米值。</param>
    /// <returns>转换后的磅值。</returns>
    float? CentimetersToPoints(float centimeters);

    /// <summary>
    /// 将测量值从毫米转换为磅（1 毫米 = 2.85 磅）。
    /// </summary>
    /// <param name="millimeters">要转换为磅的毫米值。</param>
    /// <returns>转换后的磅值。</returns>
    float? MillimetersToPoints(float millimeters);

    /// <summary>
    /// 将测量值从十二点活字转换为磅（1 十二点活字 = 12 磅）。
    /// </summary>
    /// <param name="picas">要转换为磅的十二点活字值。</param>
    /// <returns>转换后的磅值。</returns>
    float? PicasToPoints(float picas);

    /// <summary>
    /// 将测量值从行转换为磅（1 行 = 12 磅）。
    /// </summary>
    /// <param name="lines">要转换为磅的行值。</param>
    /// <returns>转换后的磅值。</returns>
    float? LinesToPoints(float lines);

    /// <summary>
    /// 将测量值从磅转换为英寸（1 英寸 = 72 磅）。
    /// </summary>
    /// <param name="points">测量值（以磅为单位）。</param>
    /// <returns>转换后的英寸值。</returns>
    float? PointsToInches(float points);

    /// <summary>
    /// 将测量值从磅转换为厘米（1 厘米 = 28.35 磅）。
    /// </summary>
    /// <param name="points">测量值（以磅为单位）。</param>
    /// <returns>转换后的厘米值。</returns>
    float? PointsToCentimeters(float points);

    /// <summary>
    /// 将测量值从磅转换为毫米（1 毫米 = 2.835 磅）。
    /// </summary>
    /// <param name="points">测量值（以磅为单位）。</param>
    /// <returns>转换后的毫米值。</returns>
    float? PointsToMillimeters(float points);

    /// <summary>
    /// 将测量值从磅转换为十二点活字（1 十二点活字 = 12 磅）。
    /// </summary>
    /// <param name="points">测量值（以磅为单位）。</param>
    /// <returns>转换后的十二点活字值。</returns>
    float? PointsToPicas(float points);

    /// <summary>
    /// 将测量值从磅转换为行（1 行 = 12 磅）。
    /// </summary>
    /// <param name="points">测量值（以磅为单位）。</param>
    /// <returns>转换后的行值。</returns>
    float? PointsToLines(float points);

    /// <summary>
    /// 将测量值从磅转换为像素。
    /// </summary>
    /// <param name="points">要转换为像素的磅值。</param>
    /// <param name="vertical">True 返回垂直像素；False 返回水平像素。</param>
    /// <returns>转换后的像素值。</returns>
    float? PointsToPixels(float points, bool? vertical = null);

    /// <summary>
    /// 将测量值从像素转换为磅。
    /// </summary>
    /// <param name="pixels">要转换为磅的像素值。</param>
    /// <param name="vertical">True 转换垂直像素；False 转换水平像素。</param>
    /// <returns>转换后的磅值。</returns>
    float? PixelsToPoints(float pixels, bool? vertical = null);

    /// <summary>
    /// 将键盘语言设置为从左到右的语言，并将文本输入方向设置为从左到右。
    /// </summary>
    void KeyboardLatin();

    /// <summary>
    /// 将键盘语言设置为从右到左的语言，并将文本输入方向设置为从右到左。
    /// </summary>
    void KeyboardBidi();

    /// <summary>
    /// 在从右到左和从左到右语言之间切换键盘语言设置。
    /// </summary>
    void ToggleKeyboard();

    /// <summary>
    /// 返回或设置键盘语言和布局设置。
    /// </summary>
    /// <param name="langId">Word 设置键盘的语言和布局组合。如果省略此参数，则返回当前语言和布局设置。</param>
    /// <returns>当前语言和布局设置。</returns>
    int? Keyboard(int langId = 0);

    /// <summary>
    /// 将 Microsoft Word 全局唯一标识符 (GUID) 作为字符串返回。
    /// </summary>
    /// <returns>Word GUID 字符串。</returns>
    string? ProductCode();

    /// <summary>
    /// 返回包含全局应用程序级别属性的 DefaultWebOptions 对象。
    /// </summary>
    /// <returns>默认 Web 选项对象。</returns>
    IWordDefaultWebOptions? DefaultWebOptions();

    /// <summary>
    /// 为 Microsoft Word 设置新文档、电子邮件或网页的默认主题。
    /// </summary>
    /// <param name="name">要分配为默认主题的主题名称加上要应用的主题格式选项。</param>
    /// <param name="documentType">为其分配默认主题的新文档类型。</param>
    void SetDefaultTheme(string name, WdDocumentMedium documentType);

    /// <summary>
    /// 返回一个字符串，表示 Microsoft Word 用于新文档、电子邮件或网页的默认主题名称加上主题格式选项。
    /// </summary>
    /// <param name="documentType">要检索默认主题名称的新文档类型。</param>
    /// <returns>默认主题名称字符串。</returns>
    string? GetDefaultTheme(WdDocumentMedium documentType);

    /// <summary>
    /// 运行 Visual Basic 宏。
    /// </summary>
    /// <param name="macroName">宏的名称。可以是模板、模块和宏名称的任何组合。</param>
    /// <param name="varg1">宏参数值。最多可以向指定宏传递 30 个参数值。</param>
    /// <param name="varg2">宏参数值。</param>
    /// <param name="varg3">宏参数值。</param>
    /// <param name="varg4">宏参数值。</param>
    /// <param name="varg5">宏参数值。</param>
    /// <param name="varg6">宏参数值。</param>
    /// <param name="varg7">宏参数值。</param>
    /// <param name="varg8">宏参数值。</param>
    /// <param name="varg9">宏参数值。</param>
    /// <param name="varg10">宏参数值。</param>
    /// <param name="varg11">宏参数值。</param>
    /// <param name="varg12">宏参数值。</param>
    /// <param name="varg13">宏参数值。</param>
    /// <param name="varg14">宏参数值。</param>
    /// <param name="varg15">宏参数值。</param>
    /// <param name="varg16">宏参数值。</param>
    /// <param name="varg17">宏参数值。</param>
    /// <param name="varg18">宏参数值。</param>
    /// <param name="varg19">宏参数值。</param>
    /// <param name="varg20">宏参数值。</param>
    /// <param name="varg21">宏参数值。</param>
    /// <param name="varg22">宏参数值。</param>
    /// <param name="varg23">宏参数值。</param>
    /// <param name="varg24">宏参数值。</param>
    /// <param name="varg25">宏参数值。</param>
    /// <param name="varg26">宏参数值。</param>
    /// <param name="varg27">宏参数值。</param>
    /// <param name="varg28">宏参数值。</param>
    /// <param name="varg29">宏参数值。</param>
    /// <param name="varg30">宏参数值。</param>
    /// <returns>宏的返回值。</returns>
    object? Run(string macroName, object? varg1 = null, object? varg2 = null, object? varg3 = null,
                object? varg4 = null, object? varg5 = null, object? varg6 = null, object? varg7 = null,
                object? varg8 = null, object? varg9 = null, object? varg10 = null, object? varg11 = null,
                object? varg12 = null, object? varg13 = null, object? varg14 = null, object? varg15 = null,
                object? varg16 = null, object? varg17 = null, object? varg18 = null, object? varg19 = null,
                object? varg20 = null, object? varg21 = null, object? varg22 = null, object? varg23 = null,
                object? varg24 = null, object? varg25 = null, object? varg26 = null, object? varg27 = null,
                object? varg28 = null, object? varg29 = null, object? varg30 = null);

    /// <summary>
    /// 打印指定文档的全部或部分内容。
    /// </summary>
    /// <param name="background">设置为 True 可在 Word 打印文档时继续运行宏。</param>
    /// <param name="append">设置为 True 可将指定文档附加到 OutputFileName 参数指定的文件名。False 覆盖 OutputFileName 的内容。</param>
    /// <param name="range">页面范围。可以是任何 WdPrintOutRange 常量。</param>
    /// <param name="outputFileName">如果 PrintToFile 为 True，此参数指定输出文件的路径和文件名。</param>
    /// <param name="from">当 Range 设置为 wdPrintFromTo 时的起始页码。</param>
    /// <param name="to">当 Range 设置为 wdPrintFromTo 时的结束页码。</param>
    /// <param name="item">要打印的项。可以是任何 WdPrintOutItem 常量。</param>
    /// <param name="copies">要打印的份数。</param>
    /// <param name="pages">要打印的页码和页面范围，用逗号分隔。例如，"2, 6-10" 打印第 2 页和第 6-10 页。</param>
    /// <param name="pageType">要打印的页面类型。可以是任何 WdPrintOutPages 常量。</param>
    /// <param name="printToFile">True 将打印机指令发送到文件。确保使用 OutputFileName 指定文件名。</param>
    /// <param name="collate">打印多份文档时，True 在打印下一份之前打印文档的所有页面。</param>
    /// <param name="fileName">要打印的文档的路径和文件名。如果省略此参数，Word 打印活动文档。</param>
    /// <param name="activePrinterMacGX">此参数仅在 Microsoft Office Macintosh 版本中可用。</param>
    /// <param name="manualDuplexPrint">True 在没有双面打印套件的打印机上打印双面文档。</param>
    /// <param name="printZoomColumn">要在一页上水平适应的页面数。可以是 1、2、3 或 4。</param>
    /// <param name="printZoomRow">要在一页上垂直适应的页面数。可以是 1、2 或 4。</param>
    /// <param name="printZoomPaperWidth">要将打印页面缩放到的宽度（以缇为单位，20 缇 = 1 磅；72 磅 = 1 英寸）。</param>
    /// <param name="printZoomPaperHeight">要将打印页面缩放到的高度（以缇为单位）。</param>
    void PrintOut(bool? background = null, bool? append = null, WdPrintOutRange? range = null,
                string? outputFileName = null, int? from = null, int? to = null,
                WdPrintOutItem? item = null, int? copies = null, string? pages = null,
                WdPrintOutPages? pageType = null, bool? printToFile = null, bool? collate = null,
                string? fileName = null, object? activePrinterMacGX = null, bool? manualDuplexPrint = null,
                int? printZoomColumn = null, int? printZoomRow = null, int? printZoomPaperWidth = null,
                int? printZoomPaperHeight = null);

    /// <summary>
    /// 加载参考书目源文件。
    /// </summary>
    /// <param name="fileName">参考书目源文件的路径和文件名。</param>
    void LoadMasterList(string fileName);

    /// <summary>
    /// 比较两个文档并返回表示包含两个文档之间差异的文档的 Document 对象，使用修订标记进行标记。
    /// </summary>
    /// <param name="originalDocument">原始文档的路径和文件名。</param>
    /// <param name="revisedDocument">要与之比较原始文档的修订文档的路径和文件名。</param>
    /// <param name="destination">指定是创建新文件还是在原始文档或修订文档中标记两个文档之间的差异。默认值为 wdCompareDestinationNew。</param>
    /// <param name="granularity">指定是按字符还是按字跟踪更改。默认值为 wdGranularityWordLevel。</param>
    /// <param name="compareFormatting">指定是否标记两个文档之间的格式差异。默认值为 True。</param>
    /// <param name="compareCaseChanges">指定是否标记两个文档之间的大小写差异。默认值为 True。</param>
    /// <param name="compareWhitespace">指定是否标记两个文档之间的空白差异，例如段落或空格。默认值为 True。</param>
    /// <param name="compareTables">指定是否比较两个文档之间表格中包含的数据差异。默认值为 True。</param>
    /// <param name="compareHeaders">指定是否比较两个文档之间的页眉和页脚差异。默认值为 True。</param>
    /// <param name="compareFootnotes">指定是否比较两个文档之间的脚注和尾注差异。默认值为 True。</param>
    /// <param name="compareTextboxes">指定是否比较两个文档之间文本框内包含的数据差异。默认值为 True。</param>
    /// <param name="compareFields">指定是否比较两个文档之间的字段差异。默认值为 True。</param>
    /// <param name="compareComments">指定是否比较两个文档之间的批注差异。默认值为 True。</param>
    /// <param name="compareMoves">指定是否比较两个文档之间的移动文本差异。默认值为 True。</param>
    /// <param name="revisedAuthor">指定比较两个文档时归因更改的人员姓名。</param>
    /// <param name="ignoreAllComparisonWarnings">指定比较两个文档时是否忽略警告。</param>
    /// <returns>包含比较结果的文档对象。</returns>
    IWordDocument? CompareDocuments(IWordDocument originalDocument, IWordDocument revisedDocument,
                WdCompareDestination destination = WdCompareDestination.wdCompareDestinationNew,
                WdGranularity granularity = WdGranularity.wdGranularityWordLevel, bool compareFormatting = true,
                bool compareCaseChanges = true, bool compareWhitespace = true, bool compareTables = true,
                bool compareHeaders = true, bool compareFootnotes = true, bool compareTextboxes = true,
                bool compareFields = true, bool compareComments = true, bool compareMoves = true,
                string revisedAuthor = "", bool ignoreAllComparisonWarnings = false);

    /// <summary>
    /// 合并两个文档并返回表示合并后包含两个文档之间差异的文档的 Document 对象，使用修订标记进行标记。
    /// </summary>
    /// <param name="originalDocument">原始文档的路径和文件名。</param>
    /// <param name="revisedDocument">要与之合并原始文档的修订文档的路径和文件名。</param>
    /// <param name="destination">指定是创建新文件还是在原始文档或修订文档中标记两个文档之间的差异。默认值为 wdCompareDestinationNew。</param>
    /// <param name="granularity">指定是按字符还是按字跟踪更改。默认值为 wdGranularityWordLevel。</param>
    /// <param name="compareFormatting">指定是否标记两个文档之间的格式差异。默认值为 True。</param>
    /// <param name="compareCaseChanges">指定是否标记两个文档之间的大小写差异。默认值为 True。</param>
    /// <param name="compareWhitespace">指定是否标记两个文档之间的空白差异，例如段落或空格。默认值为 True。</param>
    /// <param name="compareTables">指定是否比较两个文档之间表格中包含的数据差异。默认值为 True。</param>
    /// <param name="compareHeaders">指定是否比较两个文档之间的页眉和页脚差异。默认值为 True。</param>
    /// <param name="compareFootnotes">指定是否比较两个文档之间的脚注和尾注差异。默认值为 True。</param>
    /// <param name="compareTextboxes">指定是否比较两个文档之间文本框内包含的数据差异。默认值为 True。</param>
    /// <param name="compareFields">指定是否比较两个文档之间的字段差异。默认值为 True。</param>
    /// <param name="compareComments">指定是否比较两个文档之间的批注差异。默认值为 True。</param>
    /// <param name="originalAuthor">指定原始文档的作者姓名。</param>
    /// <param name="revisedAuthor">指定合并两个文档后用于未归因更改的人员姓名。</param>
    /// <param name="compareMoves">指定是否比较两个文档之间的移动文本差异。默认值为 True。</param>
    /// <param name="formatFrom">指定从哪个文档保留格式。</param>
    /// <returns>合并后的文档对象。</returns>
    IWordDocument? MergeDocuments(IWordDocument originalDocument, IWordDocument revisedDocument,
                                 WdCompareDestination destination = WdCompareDestination.wdCompareDestinationNew,
                                 WdGranularity granularity = WdGranularity.wdGranularityWordLevel, bool compareFormatting = true,
                                 bool compareCaseChanges = true, bool compareWhitespace = true,
                                 bool compareTables = true, bool compareHeaders = true, bool compareFootnotes = true,
                                 bool compareTextboxes = true, bool compareFields = true, bool compareComments = true,
                                 bool compareMoves = true, string originalAuthor = "", string revisedAuthor = "",
                                 WdMergeFormatFrom formatFrom = WdMergeFormatFrom.wdMergeFormatFromPrompt);


    /// <summary>
    /// 创建一个空白文档
    /// </summary>
    /// <returns>新建的文档对象</returns>
    [IgnoreGenerator]
    IWordDocument? BlankDocument();

    /// <summary>
    /// 从模板创建一个文档
    /// </summary>
    /// <param name="templatePath">模板路径</param>
    /// <returns>新建的文档对象</returns>
    [IgnoreGenerator]
    IWordDocument CreateFrom(string templatePath);

    /// <summary>
    /// 打开一个文档
    /// </summary>
    /// <param name="filePath">文档路径</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="password">密码</param>
    /// <returns>打开的文档对象</returns>
    [IgnoreGenerator]
    IWordDocument? Open(string filePath, bool readOnly = false, string? password = null);
}