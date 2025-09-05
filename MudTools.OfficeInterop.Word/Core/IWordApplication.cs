//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word应用程序接口
/// </summary>
public partial interface IWordApplication : IOfficeApplication
{
    IWordDocument BlankDocument();

    #region 基本属性 (Basic Properties)
    /// <summary>
    /// 获取父对象。对于 Application 对象，通常返回 null。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置活动打印机的名称。
    /// </summary>
    string ActivePrinter { get; set; }

    /// <summary>
    /// 获取表示活动文档的 Document 对象。
    /// </summary>
    IWordDocument? ActiveDocument { get; }

    /// <summary>
    /// 获取表示活动窗口的 Window 对象。
    /// </summary>
    IWordWindow? ActiveWindow { get; }

    /// <summary>
    /// 获取表示所有打开的文档的 Documents 集合。
    /// </summary>
    IWordDocuments? Documents { get; }

    /// <summary>
    /// 获取表示所有可用模板的 Templates 集合。
    /// </summary>
    IWordTemplates? Templates { get; }

    /// <summary>
    /// 获取表示所有可用加载项的 AddIns 集合。
    /// </summary>
    IWordAddIns? AddIns { get; }

    /// <summary>
    /// 获取表示 Normal 模板的 Template 对象。
    /// </summary>
    IWordTemplate? NormalTemplate { get; }

    /// <summary>
    /// 获取用于分隔文件夹名称的字符。
    /// </summary>
    string PathSeparator { get; }

    #endregion

    #region 窗口和显示属性 (Window & Display Properties)

    WordAppVisibility Visibility { get; set; }

    /// <summary>
    /// 获取或设置指定文档窗口或任务窗口的状态。
    /// </summary>
    WdWindowState WordWindowState { get; set; }

    /// <summary>
    /// 获取或设置应用程序窗口的描述文字文本。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示状态栏。
    /// </summary>
    bool DisplayStatusBar { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示滚动条。
    /// </summary>
    bool DisplayScrollBars { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在“文件”菜单上显示最近使用的文件的名称。
    /// </summary>
    bool DisplayRecentFiles { get; set; }

    /// <summary>
    /// 获取 Word 文档窗口可设置的最大宽度（以磅为单位）。
    /// </summary>
    int UsableWidth { get; }

    /// <summary>
    /// 获取 Word 文档窗口的高度设置为最大高度 (以磅为单位)。
    /// </summary>
    int UsableHeight { get; }

    /// <summary>
    /// 获取表示所有文档窗口的 Windows 集合。
    /// </summary>
    IWordWindows? Windows { get; }

    #endregion

    #region 基本方法 (Basic Methods)

    /// <summary>
    /// 退出 Microsoft Word 应用程序。
    /// </summary>
    void Quit(ref object saveChanges, ref object originalFormat, ref object routeDocument);


    /// <summary>
    /// 打印当前文档或选定内容。
    /// </summary>
    void PrintOut(ref object background, ref object append, ref object range, ref object outputFileName,
                  ref object from, ref object to, ref object item, ref object copies, ref object pages,
                  ref object pageType, ref object printToFile, ref object collate, ref object fileName,
                  ref object lineEnding, ref object outputPrinterName);

    #endregion

    #region 选择和查找属性 (Selection & Find Properties)

    /// <summary>
    /// 获取表示所选区域或插入点的 Selection 对象。
    /// </summary>
    IWordSelection? Selection { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示运行宏时的一些警告和消息的处理的方式。
    /// </summary>
    WdAlertLevel DisplayAlerts { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时显示自动完成提示。
    /// </summary>
    bool DisplayAutoCompleteTips { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将批注、脚注、尾注和超链接显示为提示。
    /// </summary>
    bool DisplayScreenTips { get; set; }

    #endregion

    #region 选项和设置属性 (Options & Settings Properties)

    /// <summary>
    /// 获取表示 Microsoft Word 中应用程序设置的 Options 对象。
    /// </summary>
    IWordOptions? Options { get; } // 假设 Options 未封装

    /// <summary>
    /// 获取或设置一个值，该值指示 Word 处理 Ctrl+Break 用户中断的方式。
    /// </summary>
    WdEnableCancelKey EnableCancelKey { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 如何处理调用需要尚未安装的功能的方法和属性。
    /// </summary>
    MsoFeatureInstall FeatureInstall { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 在键入时是否自动检测所使用的语言。
    /// </summary>
    bool CheckLanguage { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否打开屏幕更新。
    /// </summary>
    bool ScreenUpdating { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否打开拼写和语法检查。
    /// </summary>
    bool CheckSpellingAsYouType { get; set; }
    bool CheckGrammarAsYouType { get; set; }

    #endregion

    #region 语言和字典属性 (Language & Dictionary Properties)

    /// <summary>
    /// 获取表示“语言”对话框中列出的校对语言的 Languages 集合。
    /// </summary>
    IWordLanguages? Languages { get; }

    /// <summary>
    /// 获取表示所有可用字体名称的 FontNames 集合。
    /// </summary>
    IWordFontNames? FontNames { get; }

    /// <summary>
    /// 获取表示所有可用纵向字体名称的 FontNames 集合。
    /// </summary>
    IWordFontNames? PortraitFontNames { get; }

    /// <summary>
    /// 获取表示所有可用横向字体名称的 FontNames 集合。
    /// </summary>
    IWordFontNames? LandscapeFontNames { get; }

    /// <summary>
    /// 获取表示活动自定义字典集合的 Dictionaries 对象。
    /// </summary>
    IWordDictionaries? CustomDictionaries { get; }

    #endregion

    #region 自动更正和列表属性 (AutoCorrect & Lists Properties)

    /// <summary>
    /// 获取表示当前自动更正选项、条目和异常的 AutoCorrect 对象。
    /// </summary>
    IWordAutoCorrect? AutoCorrect { get; }

    /// <summary>
    /// 获取表示对电子邮件进行的自动更正的 AutoCorrect 对象。
    /// </summary>
    IWordAutoCorrect? AutoCorrectEmail { get; }

    /// <summary>
    /// 获取表示项目符号、编号和大纲编号模板库的 ListGalleries 集合。
    /// </summary>
    IWordListGalleries? ListGalleries { get; }

    #endregion

    #region 文件和模板属性 (File & Template Properties)

    /// <summary>
    /// 获取表示最近访问的文件的 RecentFiles 集合。
    /// </summary>
    IWordRecentFiles RecentFiles { get; }

    /// <summary>
    /// 获取或设置启动文件夹的完整路径（不包括最后的分隔符）。
    /// </summary>
    string StartupPath { get; set; }

    /// <summary>
    /// 获取或设置用户的邮件地址。
    /// </summary>
    string UserAddress { get; set; }

    /// <summary>
    /// 获取或设置用户的姓名缩写。
    /// </summary>
    string UserInitials { get; set; }

    /// <summary>
    /// 获取或设置用户名。
    /// </summary>
    string UserName { get; set; }

    #endregion

    #region 更多方法 (More Methods)

    /// <summary>
    /// 打开一个现有文档。
    /// </summary>
    /// <param name="fileName">要打开的文档的文件名。</param>
    /// <param name="confirmConversions">如果为 true，则在文件不是 Word 格式时显示“转换文件”对话框。</param>
    /// <param name="readOnly">如果为 true，则以只读方式打开文档。</param>
    /// <param name="addToRecentFiles">如果为 true，则将文件添加到最近使用的文件列表中。</param>
    /// <param name="passwordDocument">打开文档所需的密码。</param>
    /// <param name="passwordTemplate">打开模板所需的密码。</param>
    /// <param name="revert">如果为 true，则将文档恢复到上次保存的版本。</param>
    /// <param name="writePasswordDocument">保存对文档所做的更改所需的密码。</param>
    /// <param name="writePasswordTemplate">保存对模板所做的更改所需的密码。</param>
    /// <param name="format">文档的格式。</param>
    /// <param name="encoding">文档的编码。</param>
    /// <param name="visible">如果为 true，则打开文档时使其可见。</param>
    /// <returns>打开的文档对象。</returns>
    IWordDocument? OpenDocument(string fileName, bool confirmConversions = true, bool readOnly = false, bool addToRecentFiles = true,
                                     string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "",
                                     string writePasswordTemplate = "", WdOpenFormat format = WdOpenFormat.wdOpenFormatAuto,
                                     MsoEncoding encoding = MsoEncoding.msoEncodingSimplifiedChineseAutoDetect, bool visible = true);

    /// <summary>
    /// 新建一个文档。
    /// </summary>
    /// <param name="template">用于创建新文档的模板。</param>
    /// <param name="newTemplate">如果为 true，则将文档创建为新模板。</param>
    /// <returns>新建的文档对象。</returns>
    IWordDocument? NewDocument(object template, object newTemplate);

    /// <summary>
    /// 执行查找操作。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <returns>如果找到则返回 true，否则返回 false。</returns>
    bool FindText(string findText);

    /// <summary>
    /// 替换文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceWith">替换文本。</param>
    /// <param name="replace">替换操作类型。</param>
    /// <returns>替换的次数。</returns>
    int ReplaceText(string findText, string replaceWith, MsWord.WdReplace replace);

    /// <summary>
    /// 使 Visual Basic 编辑器窗口可见或不可见。
    /// </summary>
    /// <param name="visible">如果为 true，则使窗口可见。</param>
    void ShowVisualBasicEditor(bool visible);

    /// <summary>
    /// 获取有关当前国家/地区和国际设置的信息。
    /// </summary>
    /// <param name="index">要返回的信息类型。</param>
    /// <returns>相关信息。</returns>
    object GetInternational(WdInternationalIndex index);

    #endregion

    #region 自动化和 COM 属性 (Automation & COM Properties)
    /// <summary>
    /// 获取一个值，该值指示引用对象的指定变量是否有效。
    /// </summary>
    /// <param name="obj">要检查的对象。</param>
    /// <returns>如果对象有效则返回 true，否则返回 false。</returns>
    bool IsObjectValid(object obj);

    /// <summary>
    /// 获取表示所有可用文件转换器的 FileConverters 集合。
    /// </summary>
    IWordFileConverters? FileConverters { get; }

    /// <summary>
    /// 获取表示所有正在运行的应用程序的 Tasks 集合。
    /// </summary>
    IWordTasks? Tasks { get; }

    /// <summary>
    /// 获取表示所有内置对话框的 Dialogs 集合。
    /// </summary>
    IWordDialogs? Dialogs { get; }

    /// <summary>
    /// 获取表示所有自定义键绑定的 KeyBindings 集合。
    /// </summary>
    IWordKeyBindings? KeyBindings { get; }

    /// <summary>
    /// 获取表示所有已加载的 COM 加载项的 COMAddIns 集合。
    /// </summary>
    object COMAddIns { get; }

    #endregion

    #region 邮件相关属性 (Mail Properties)

    /// <summary>
    /// 获取表示电子邮件创作的全局首选项的 EmailOptions 对象。
    /// </summary>
    IWordEmailOptions? EmailOptions { get; }

    /// <summary>
    /// 获取或设置用于电子邮件的模板。
    /// </summary>
    string EmailTemplate { get; set; }

    /// <summary>
    /// 获取表示邮件标签的 MailingLabel 对象。
    /// </summary>
    IWordMailingLabel? MailingLabel { get; }

    /// <summary>
    /// 获取表示活动电子邮件的 MailMessage 对象。
    /// </summary>
    IWordMailMessage MailMessage { get; }

    /// <summary>
    /// 获取邮件系统的类型。
    /// </summary>
    WdMailSystem MailSystem { get; }

    /// <summary>
    /// 获取一个值，该值指示是否安装了 MAPI。
    /// </summary>
    bool MAPIAvailable { get; }

    /// <summary>
    /// 获取一个值，该值指示插入点是否位于电子邮件标头字段中。
    /// </summary>
    bool FocusInMailHeader { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在全屏模式下打开附件。
    /// </summary>
    bool OpenAttachmentsInFullScreen { get; set; }

    #endregion

    #region 安全性相关属性 (Security Properties)

    /// <summary>
    /// 获取或设置自动化安全级别。
    /// </summary>
    MsoAutomationSecurity AutomationSecurity { get; set; }

    /// <summary>
    /// 获取或设置文件验证方式。
    /// </summary>
    MsoFileValidationMode FileValidation { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否限制链接样式。
    /// </summary>
    bool RestrictLinkedStyles { get; set; }

    #endregion

    #region 系统和环境属性 (System & Environment Properties)
    /// <summary>
    /// 获取一个值，该值指示是否安装了数学协处理器。
    /// </summary>
    bool MathCoprocessorAvailable { get; }

    /// <summary>
    /// 获取一个值，该值指示是否有可用于系统的鼠标。
    /// </summary>
    bool MouseAvailable { get; }

    /// <summary>
    /// 获取 NUM LOCK 键的状态。
    /// </summary>
    bool NumLock { get; }

    /// <summary>
    /// 获取 CAPS LOCK 键的状态。
    /// </summary>
    bool CapsLock { get; }

    /// <summary>
    /// 获取一个值，该值指示文档或应用程序是否由用户创建或打开。
    /// </summary>
    bool UserControl { get; }

    #endregion

    #region 更多方法 (More Methods)

    /// <summary>
    /// 保护文档。
    /// </summary>
    /// <param name="type">保护类型。</param>
    /// <param name="noReset">是否重置。</param>
    /// <param name="password">密码。</param>
    /// <param name="useIRM">是否使用信息权限管理。</param>
    /// <param name="enforceStyleLock">是否强制样式锁定。</param>
    void Protect(MsWord.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock);

    /// <summary>
    /// 取消保护文档。
    /// </summary>
    /// <param name="password">密码。</param>
    void Unprotect(object password);

    /// <summary>
    /// 保存所有打开的文档。
    /// </summary>
    void SaveAll();

    /// <summary>
    /// 获取指定键绑定。
    /// </summary>
    /// <param name="keyCode">键代码。</param>
    /// <param name="keyCode2">第二个键代码（可选）。</param>
    /// <returns>键绑定对象。</returns>
    IWordKeyBinding FindKey(int keyCode, object keyCode2);

    /// <summary>
    /// 获取分配给指定项的所有组合键。
    /// </summary>
    /// <param name="keyCategory">键类别。</param>
    /// <param name="command">命令。</param>
    /// <param name="commandParameter">命令参数。</param>
    /// <returns>键绑定集合。</returns>
    IWordKeysBoundTo KeysBoundTo(MsWord.WdKeyCategory keyCategory, string command, object commandParameter);

    /// <summary>
    /// 获取同义词信息。
    /// </summary>
    /// <param name="word">要查询的单词。</param>
    /// <param name="languageID">语言ID。</param>
    /// <returns>同义词信息对象。</returns>
    IWordSynonymInfo SynonymInfo(string word, object languageID);

    /// <summary>
    /// 获取文件对话框。
    /// </summary>
    /// <param name="fileDialogType">文件对话框类型。</param>
    /// <returns>文件对话框对象。</returns>
    IOfficeFileDialog FileDialog(MsoFileDialogType fileDialogType);

    /// <summary>
    /// 获取智能标记识别器集合。
    /// </summary>
    IWordSmartTagRecognizers SmartTagRecognizers { get; }

    /// <summary>
    /// 获取智能标记类型集合。
    /// </summary>
    IWordSmartTagTypes SmartTagTypes { get; }

    #endregion

    #region 剩余属性 (Remaining Properties)

    /// <summary>
    /// 获取表示 AnswerWizard 对象，其中包含联机帮助搜索引擎使用的文件。
    /// </summary>
    MsWord.AnswerWizard AnswerWizard { get; }

    /// <summary>
    /// 获取一个值，该值指示是否支持任意 XML。
    /// </summary>
    bool ArbitraryXMLSupportAvailable { get; }

    /// <summary>
    /// 获取表示 Microsoft Office 帮助查看器的 IAssistance 对象。
    /// </summary>
    IOfficeAssistance? Assistance { get; }

    /// <summary>
    /// 获取表示在将表格和图片等项目插入文档中时自动添加的标题的 AutoCaptions 集合。
    /// </summary>
    IWordAutoCaptions? AutoCaptions { get; }

    /// <summary>
    /// 获取后台打印队列中打印作业的编号。
    /// </summary>
    int BackgroundPrintingStatus { get; }

    /// <summary>
    /// 获取排队在后台保存的文件数。
    /// </summary>
    int BackgroundSavingStatus { get; }

    /// <summary>
    /// 获取表示 Microsoft Office Word 中存储的书目引用源的 Bibliography 对象。
    /// </summary>
    IWordBibliography Bibliography { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否可以使用 Word 打开 HTML 文件。
    /// </summary>
    string BrowseExtraFileTypes { get; set; }

    /// <summary>
    /// 获取表示垂直滚动条上的“选择浏览对象”工具的 Browser 对象。
    /// </summary>
    IWordBrowser Browser { get; }

    /// <summary>
    /// 获取 Word 应用程序的内部版本号。
    /// </summary>
    string BuildFull { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示在比较和合并文档时是否默认使用“法律黑线”选项。
    /// </summary>
    bool DefaultLegalBlackline { get; set; }

    /// <summary>
    /// 获取或设置在“另存为”对话框中的“另存为类型”框中显示的默认格式。
    /// </summary>
    string DefaultSaveFormat { get; set; }

    /// <summary>
    /// 获取或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。
    /// </summary>
    string DefaultTableSeparator { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示文档信息面板。
    /// </summary>
    bool DisplayDocumentInformationPanel { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在全屏模式下打开附件。
    /// </summary>
    bool DontResetInsertionPointProperties { get; set; }

    /// <summary>
    /// 获取表示所有活动的自定义转换字典的 HangulHanjaConversionDictionaries 集合。
    /// </summary>
    MsWord.HangulHanjaConversionDictionaries HangulHanjaDictionaries { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在受保护的视图中打开文件。
    /// </summary>
    bool IsSandboxed { get; }

    /// <summary>
    /// 获取表示所选 Microsoft Word 用户界面的语言设置。
    /// </summary>
    MsoLanguageID Language { get; }

    /// <summary>
    /// 获取表示有关 Microsoft Word 中的语言设置的信息的 LanguageSettings 对象。
    /// </summary>
    MsWord.LanguageSettings LanguageSettings { get; }

    /// <summary>
    /// 获取表示在其中存储包含正在运行过程的模块的模板或文档的 Template 或 Document 对象。
    /// </summary>
    object MacroContainer { get; } // 使用 object 以避免依赖

    /// <summary>
    /// 获取表示公式的自动更正条目的 OMathAutoCorrect 对象。
    /// </summary>
    IWordOMathAutoCorrect OMathAutoCorrect { get; }

    /// <summary>
    /// 获取一个 PickerDialog 对象，该对象提供在对话框中选择人员或数据的功能。
    /// </summary>
    object PickerDialog { get; } // 使用 object 以避免依赖

    /// <summary>
    /// 获取一个值，该值指示打印预览是否为当前视图。
    /// </summary>
    bool PrintPreview { get; set; }

    /// <summary>
    /// 获取表示所有受保护的视图窗口的 ProtectedViewWindows 集合。
    /// </summary>
    IWordProtectedViewWindows ProtectedViewWindows { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示启动 Microsoft Word 时是否显示任务窗格。
    /// </summary>
    bool ShowStartupDialog { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示样式预览。
    /// </summary>
    bool ShowStylePreviews { get; set; }

    /// <summary>
    /// 获取一个值，该值指示应用程序是否处于特殊模式（例如 CopyText 模式或 MoveText 模式）。
    /// </summary>
    bool SpecialMode { get; }

    /// <summary>
    /// 获取或设置状态栏中显示的文本。
    /// </summary>
    string StatusBar { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Word 中最常执行的任务的 TaskPanes 集合。
    /// </summary>
    IWordTaskPanes? TaskPanes { get; }

    /// <summary>
    /// 获取一个 UndoRecord 对象，该对象提供撤消堆栈中的自定义入口点。
    /// </summary>
    object UndoRecord { get; }

    /// <summary>
    /// 获取自动化对象 (Word.Basic) ，其中包括 Microsoft Word 6.0 版和 Windows 95 Word 中提供的所有 WordBasic 语句和函数的方法。
    /// </summary>
    object WordBasic { get; }

    /// <summary>
    /// 获取表示当前在应用程序中加载的一组颜色样式的 SmartArtColors 对象。
    /// </summary>
    object SmartArtColors { get; }

    /// <summary>
    /// 获取表示当前在应用程序中加载的 SmartArt 布局集的 SmartArtLayouts 对象。
    /// </summary>
    object SmartArtLayouts { get; }

    /// <summary>
    /// 获取表示应用程序中当前加载的 SmartArt 样式集的 SmartArtQuickStyles 对象。
    /// </summary>
    object SmartArtQuickStyles { get; }

    /// <summary>
    /// 获取表示活动加密会话的对象。
    /// </summary>
    object ActiveEncryptionSession { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示图表数据点是否被跟踪。
    /// </summary>
    bool ChartDataPointTrack { get; set; }

    /// <summary>
    /// 获取一个 FileSearch 对象，该对象可用于使用绝对路径或相对路径搜索文件。
    /// </summary>
    IWordFileSearch FileSearch { get; }

    #endregion

    #region 剩余方法 (Remaining Methods)

    /// <summary>
    /// 将文档另存为 PDF 或 XPS 格式。
    /// </summary>
    /// <param name="outputFileName">文件名。</param>
    /// <param name="exportFormat">导出格式。</param>
    /// <param name="openAfterExport">导出后是否打开文件。</param>
    /// <param name="optimizeFor">优化目标。</param>
    /// <param name="item">要导出的项。</param>
    /// <param name="includeDocProps">是否包含文档属性。</param>
    /// <param name="keepIRM">是否保留 IRM。</param>
    /// <param name="createBookmarks">创建书签的方式。</param>
    /// <param name="docStructureTags">是否包含文档结构标签。</param>
    /// <param name="bitmapMissingFonts">是否将缺失的字体作为位图嵌入。</param>
    /// <param name="useISO19005_1">是否使用 ISO 19005-1 (PDF/A)。</param>
    /// <param name="fixedFormatExtClassPtr">固定格式意图。</param>
    void ExportAsFixedFormat(string outputFileName,
        WdExportFormat exportFormat,
        bool openAfterExport = false,
        WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
        WdExportRange range = WdExportRange.wdExportAllDocument,
        int from = 1, int to = 1,
        WdExportItem item = WdExportItem.wdExportDocumentContent,
        bool includeDocProps = false,
        bool keepIRM = true,
        WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
        bool docStructureTags = true,
        bool bitmapMissingFonts = true,
        bool useISO19005_1 = false,
         object fixedFormatExtClassPtr = null);
    #endregion
}
