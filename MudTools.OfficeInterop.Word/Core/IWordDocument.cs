//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 文档接口，用于操作 Word 文档
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDocument : IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取文档完整路径
    /// </summary>
    string FullName { get; }

    /// <summary>
    /// 获取或设置文档的加密提供程序
    /// </summary>
    string EncryptionProvider { get; set; }

    /// <summary>
    /// 获取或设置文档内置属性集合
    /// </summary>
    IOfficeDocumentProperties? BuiltInDocumentProperties { get; }

    /// <summary>
    /// 获取或设置文档作者
    /// </summary>
    [IgnoreGenerator]
    string Author { get; set; }

    /// <summary>
    /// 获取或设置文档主题
    /// </summary>
    [IgnoreGenerator]
    string Subject { get; set; }

    /// <summary>
    /// 获取或设置文档描述
    /// </summary>
    [IgnoreGenerator]
    string Description { get; set; }

    /// <summary>
    /// 获取或设置文档关键字
    /// </summary>
    [IgnoreGenerator]
    string Keywords { get; set; }

    /// <summary>
    /// 获取或设置文档公司信息
    /// </summary>
    [IgnoreGenerator]
    string Company { get; set; }

    /// <summary>
    /// 获取或设置文档标题
    /// </summary>
    [IgnoreGenerator]
    string Title { get; set; }

    /// <summary>
    /// 获取文档路径
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取文档是否已修改
    /// </summary>
    bool? Saved { get; set; }

    /// <summary>
    /// 获取文档是否已发送路由
    /// </summary>
    bool? Routed { get; }

    /// <summary>
    /// 获取文档是否为主控文档
    /// </summary>
    bool? IsMasterDocument { get; }

    /// <summary>
    /// 获取或设置是否自动断字
    /// </summary>
    bool? AutoHyphenation { get; set; }

    /// <summary>
    /// 获取或设置是否嵌入 TrueType 字体
    /// </summary>
    bool? EmbedTrueTypeFonts { get; set; }

    /// <summary>
    /// 获取或设置是否保存窗体数据
    /// </summary>
    bool? SaveFormsData { get; set; }

    /// <summary>
    /// 获取文档是否为子文档
    /// </summary>
    bool? IsSubdocument { get; }

    /// <summary>
    /// 获取文档保存格式
    /// </summary>
    int? SaveFormat { get; }

    /// <summary>
    /// 获取或设置是否保存子集字体
    /// </summary>
    bool? SaveSubsetFonts { get; set; }

    /// <summary>
    /// 获取或设置是否仅打印窗体数据到预打印的表单上
    /// </summary>
    bool? PrintFormsData { get; set; }

    /// <summary>
    /// 获取或设置是否显示语法错误
    /// </summary>
    bool? ShowGrammaticalErrors { get; set; }

    /// <summary>
    /// 获取或设置是否已完成拼写检查
    /// </summary>
    bool? SpellingChecked { get; set; }

    /// <summary>
    /// 获取或设置是否显示摘要
    /// </summary>
    bool? ShowSummary { get; set; }

    /// <summary>
    /// 获取或设置是否显示拼写错误
    /// </summary>
    bool? ShowSpellingErrors { get; set; }

    /// <summary>
    /// 获取或设置是否已完成语法检查
    /// </summary>
    bool? GrammarChecked { get; set; }

    /// <summary>
    /// 获取或设置是否打印小数宽度字符
    /// </summary>
    bool? PrintFractionalWidths { get; set; }

    /// <summary>
    /// 获取或设置是否在文本上打印 PostScript
    /// </summary>
    bool? PrintPostScriptOverText { get; set; }

    /// <summary>
    /// 获取或设置打开文档时是否更新样式
    /// </summary>
    bool? UpdateStylesOnOpen { get; set; }

    /// <summary>
    /// 获取或设置是否建议以只读方式打开文档
    /// </summary>
    bool? ReadOnlyRecommended { get; set; }

    /// <summary>
    /// 获取或设置是否对大写字母进行断字
    /// </summary>
    bool? HyphenateCaps { get; set; }

    /// <summary>
    /// 获取或设置断字区域宽度（单位：磅）
    /// </summary>
    int? HyphenationZone { get; set; }

    /// <summary>
    /// 获取或设置摘要长度
    /// </summary>
    int? SummaryLength { get; set; }

    /// <summary>
    /// 获取或设置默认制表位宽度（单位：磅）
    /// </summary>
    float? DefaultTabStop { get; set; }

    /// <summary>
    /// 获取或设置连续断字符的最大数量
    /// </summary>
    int? ConsecutiveHyphensLimit { get; set; }

    /// <summary>
    /// 获取或设置文档是否有路由传阅单
    /// </summary>
    bool? HasRoutingSlip { get; set; }

    /// <summary>
    /// 获取或设置文档类型
    /// </summary>
    WdDocumentKind Kind { get; set; }

    /// <summary>
    /// 获取文档的权限设置对象，用于管理和控制文档的访问权限
    /// </summary>
    IOfficePermission? Permission { get; }

    /// <summary>
    /// 获取文档是否受保护
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 获取文档保护类型
    /// </summary>
    WdProtectionType ProtectionType { get; }

    /// <summary>
    /// 获取文档类型
    /// </summary>
    WdDocumentType Type { get; }

    /// <summary>
    /// 获取文档的命令栏集合
    /// </summary>
    IOfficeCommandBars? CommandBars { get; }

    /// <summary>
    /// 获取文档的批注集合
    /// </summary>
    IWordComments? Comments { get; }

    /// <summary>
    /// 获取文档的尾注集合
    /// </summary>
    IWordEndnotes? Endnotes { get; }

    /// <summary>
    /// 获取文档的脚注集合
    /// </summary>
    IWordFootnotes? Footnotes { get; }

    /// <summary>
    /// 获取文档的单词集合
    /// </summary>
    IWordWords? Words { get; }

    /// <summary>
    /// 获取文档的主要内容范围
    /// </summary>
    IWordRange? Content { get; }

    /// <summary>
    /// 获取文档的字符集合
    /// </summary>
    IWordCharacters? Characters { get; }

    /// <summary>
    /// 获取文档的域集合
    /// </summary>
    IWordFields? Fields { get; }

    /// <summary>
    /// 获取文档的窗体域集合
    /// </summary>
    IWordFormFields? FormFields { get; }

    /// <summary>
    /// 获取文档的目录集合
    /// </summary>
    IWordTablesOfContents? TablesOfContents { get; }

    /// <summary>
    /// 获取文档的引文目录集合
    /// </summary>
    IWordTablesOfAuthorities? TablesOfAuthorities { get; }

    /// <summary>
    /// 获取文档的框架集合
    /// </summary>
    IWordFrames? Frames { get; }

    /// <summary>
    /// 获取或设置文档的页面设置
    /// </summary>
    IWordPageSetup? PageSetup { get; }

    /// <summary>
    /// 获取文档的窗口集合
    /// </summary>
    IWordWindows? Windows { get; }

    /// <summary>
    /// 获取文档的信封对象，用于操作文档中的信封相关内容
    /// </summary>
    IWordEnvelope? Envelope { get; }

    /// <summary>
    /// 获取Word 文档的邮件合并功能的二次封装接口。
    /// </summary>
    IWordMailMerge? MailMerge { get; }

    /// <summary>
    /// 获取或设置文档的背景形状
    /// </summary>
    IWordShape? Background { get; }

    /// <summary>
    /// 获取文档页数
    /// </summary>
    [IgnoreGenerator]
    int PageCount { get; }

    /// <summary>
    /// 获取文档中的内嵌形状集合。
    /// 内嵌形状是嵌入在文本行中的对象，如图片、图表或OLE对象，它们随着文本移动而移动。
    /// </summary>
    IWordInlineShapes? InlineShapes { get; }

    /// <summary>
    /// 获取文档中的浮动形状集合。
    /// 浮动形状是独立于文本流的对象，可以放置在页面上的任意位置，并可以设置文字环绕方式。
    /// </summary>
    IWordShapes? Shapes { get; }

    /// <summary>
    /// 获取活动窗口
    /// </summary>
    IWordWindow? ActiveWindow { get; }

    /// <summary>
    /// 获取文档范围集合
    /// </summary>
    IWordStoryRanges? StoryRanges { get; }

    /// <summary>
    /// 获取文档书签集合
    /// </summary>
    IWordBookmarks? Bookmarks { get; }

    /// <summary>
    /// 获取文档表格集合
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取文档段落集合
    /// </summary>
    IWordParagraphs? Paragraphs { get; }

    /// <summary>
    /// 获取文档节集合
    /// </summary>
    IWordSections? Sections { get; }

    /// <summary>
    /// 获取文档样式集合
    /// </summary>
    IWordStyles? Styles { get; }

    /// <summary>
    /// 获取文档列表模板集合
    /// </summary>
    IWordListTemplates? ListTemplates { get; }

    /// <summary>
    /// 获取文档变量集合
    /// </summary>
    IWordVariables? Variables { get; }


    IWordOMaths? OMaths { get; }

    WdOMathBreakBin OMathBreakBin { get; set; }

    WdOMathBreakSub OMathBreakSub { get; set; }

    WdOMathJc OMathJc { get; set; }

    float OMathLeftMargin { get; set; }

    float OMathRightMargin { get; set; }

    float OMathWrap { get; set; }

    bool OMathIntSubSupLim { get; set; }

    bool OMathNarySupSubLim { get; set; }

    bool OMathSmallFrac { get; set; }

    string OMathFontName { get; set; }


    bool UseMathDefaults { get; set; }

    IWordContentControls? ContentControls { get; }

    IWordBibliography? Bibliography { get; }


    IOfficeScripts? Scripts { get; }

    IWordXMLNodes? XMLSchemaViolations { get; }

    IWordXMLNodes? XMLNodes { get; }

    IWordXMLChildNodeSuggestions? ChildNodeSuggestions { get; }


    IWordHTMLDivisions? HTMLDivisions { get; }

    IWordSmartTags? SmartTags { get; }

    bool LockTheme { get; set; }

    bool LockQuickStyleSet { get; set; }

    string OriginalDocumentTitle { get; }

    string RevisedDocumentTitle { get; }

    bool FormattingShowNextLevel { get; set; }

    bool FormattingShowUserStyleName { get; set; }

    bool Final { get; set; }

    bool HasVBProject { get; }

    int DocID { get; }

    int CurrentRsid { get; }

    string WordOpenXML { get; }

    int CompatibilityMode { get; }

    bool ChartDataPointTrack { get; set; }

    bool IsInAutosave { get; }

    bool TrackFormatting { get; set; }

    bool TrackMoves { get; set; }

    WdStyleSort StyleSortMethod { get; set; }

    int ReadingLayoutSizeY { get; set; }

    int ReadingLayoutSizeX { get; set; }

    bool RemoveDateAndTime { get; set; }

    bool ReadingModeLayoutFrozen { get; set; }

    string XMLSaveThroughXSLT { get; set; }

    bool XMLUseXSLTWhenSaving { get; set; }

    bool XMLShowAdvancedErrors { get; set; }

    bool XMLHideNamespaces { get; set; }

    bool XMLSaveDataOnly { get; set; }

    bool AutoFormatOverride { get; set; }

    bool EnforceStyle { get; set; }

    WdShowFilter FormattingShowFilter { get; set; }

    bool FormattingShowNumbering { get; set; }

    bool FormattingShowParagraph { get; set; }

    bool FormattingShowClear { get; set; }

    bool FormattingShowFont { get; set; }

    bool EmbedLinguisticData { get; set; }

    bool PasswordEncryptionFileProperties { get; }

    int PasswordEncryptionKeyLength { get; }

    string PasswordEncryptionAlgorithm { get; }

    string PasswordEncryptionProvider { get; }

    object DefaultTableStyle { get; }

    WdLineEndingType TextLineEnding { get; set; }

    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding TextEncoding { get; set; }


    bool SmartTagsAsXMLProps { get; set; }

    bool EmbedSmartTags { get; set; }

    bool RemovePersonalInformation { get; set; }

    WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter { get; set; }

    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding SaveEncoding { get; set; }

    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding OpenEncoding { get; }

    string DefaultTargetFrame { get; set; }

    bool DoNotEmbedSystemFonts { get; set; }

    bool DisableFeatures { get; set; }

    bool VBASigned { get; }

    [ComPropertyWrap(IsMethod = true)]
    object ClickAndTypeParagraphStyle { get; set; }

    bool LanguageDetected { get; set; }

    string ActiveThemeDisplayName { get; }

    string ActiveTheme { get; }

    /// <summary>
    /// 获取或设置文档密码
    /// </summary>
    string Password { set; }

    /// <summary>
    /// 获取文档是否设置了密码保护
    /// </summary>
    bool HasPassword { get; }

    /// <summary>
    /// 获取或设置文档写保护密码
    /// </summary>
    string WritePassword { set; }

    /// <summary>
    /// 获取文档的统计信息，如页数、字数、字符数等
    /// </summary>
    /// <param name="Statistic">指定要计算的统计信息类型</param>
    /// <param name="IncludeFootnotesAndEndnotes">是否包含脚注和尾注，默认为 null</param>
    /// <returns>指定统计信息的数值</returns>
    int? ComputeStatistics(WdStatistic Statistic, bool? IncludeFootnotesAndEndnotes = null);

    /// <summary>
    /// 激活文档
    /// </summary>
    void Activate();

    /// <summary>
    /// 保存文档
    /// </summary>
    void Save();

    void SaveAs(string? fileName = null, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                string? password = null, bool? addToRecentFiles = null, string? writePassword = null,
                bool? readOnlyRecommended = null, bool? embedTrueTypeFonts = null,
                bool? saveNativePictureFormat = null, bool? saveFormsData = null, bool? saveAsAOCELetter = null,
                [ComNamespace("MsCore")] MsoEncoding? encoding = null, bool? insertLineBreaks = null, bool? allowSubstitutions = null,
                WdLineEndingType? lineEnding = null, bool? addBiDiMarks = null);

    void SaveAs2(string fileName, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                   string? password = null, bool? addToRecentFiles = null, string? writePassword = null,
                   bool? readOnlyRecommended = null, bool? embedTrueTypeFonts = null,
                   bool? saveNativePictureFormat = null, bool? saveFormsData = null,
                   bool? saveAsAOCELetter = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null, bool? insertLineBreaks = null,
                   bool? allowSubstitutions = null, WdLineEndingType? lineEnding = null,
                   bool? addBiDiMarks = null, WdCompatibilityMode? compatibilityMode = null);

    void SaveCopyAs(string fileName, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                    string? password = null, bool? addToRecentFiles = null, bool? writePassword = null, bool? readOnlyRecommended = null,
                    bool? embedTrueTypeFonts = null, bool? saveNativePictureFormat = null, bool? saveFormsData = null,
                    bool? saveAsAOCELetter = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null,
                    bool? insertLineBreaks = null, bool? allowSubstitutions = null,
                    bool? lineEnding = null, bool? addBiDiMarks = null, WdCompatibilityMode? compatibilityMode = null);



    /// <summary>
    /// 另存为文档
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="fileFormat">文件格式</param>
    void SaveAs(string fileName, WdSaveFormat fileFormat = WdSaveFormat.wdFormatDocumentDefault);

    /// <summary>
    /// 关闭当前文档
    /// </summary>
    void Close(WdSaveOptions? saveOptions = null, WdOriginalFormat? originalFormat = null, bool? routeDocument = null);

    /// <summary>
    /// 打印文档
    /// </summary>
    /// <param name="background">如果为 true，则将文档打印到后台打印机队列</param>
    /// <param name="append">如果为 true，则将指定的文档追加到活动打印机的当前打印作业中</param>
    /// <param name="range">要打印的文档部分</param>
    /// <param name="outputFileName">要将输出发送到的文件的路径和名称</param>
    /// <param name="item">要打印的项目</param>
    /// <param name="copies">要打印的份数</param>
    /// <param name="pages">要打印的页码范围</param>
    /// <param name="pageType">要打印的页面类型（所有页面、奇数页或偶数页）</param>
    /// <param name="printToFile">如果为 true，则将打印输出发送到文件</param>
    /// <param name="collate">如果为 true，则对多份打印进行校对</param>
    /// <param name="manualDuplexPrint">如果为 true，则在手动双面打印机上打印文档</param>
    /// <param name="printZoomColumn">水平打印的页数</param>
    /// <param name="printZoomRow">垂直打印的页数</param>
    /// <param name="printZoomPaperWidth">缩放页的宽度（百分比）</param>
    /// <param name="printZoomPaperHeight">缩放页的高度（百分比）</param>
    void PrintOut(bool? background = null,
        bool? append = null, WdPrintOutRange? range = null,
        string? outputFileName = null,
        WdPrintOutItem? item = null, int? copies = null, string? pages = null,
        WdPrintOutPages? pageType = null, bool? printToFile = null,
        bool? collate = null, bool? manualDuplexPrint = null,
        int? printZoomColumn = null, int? printZoomRow = null,
        int? printZoomPaperWidth = null, int? printZoomPaperHeight = null);

    /// <summary>
    /// 打印文档
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="pages">打印页码范围</param>
    void PrintOut(int copies, string pages = "");

    /// <summary>
    /// 保护文档
    /// </summary>
    /// <param name="protectionType">保护类型</param>
    /// <param name="password">密码（可选）</param>
    /// <param name="noReset">是否不重置现有保护</param>
    void Protect(WdProtectionType protectionType, string? password = null, bool? noReset = null);

    /// <summary>
    /// 取消文档保护
    /// </summary>
    /// <param name="password">密码（可选）</param>
    void Unprotect(string? password = null);


    /// <summary>
    /// 接受所有修订
    /// </summary>
    void AcceptAllRevisions();

    /// <summary>
    /// 拒绝所有修订
    /// </summary>
    void RejectAllRevisions();

    /// <summary>
    /// 选择所有可编辑区域
    /// </summary>
    /// <param name="editorID">编辑器标识符，如果为null则表示选择所有编辑器的可编辑区域</param>
    void SelectAllEditableRanges(string? editorID = null);

    /// <summary>
    /// 选择所有可编辑区域
    /// </summary>
    /// <param name="editorID">编辑器类型</param>
    void SelectAllEditableRanges(WdEditorType editorID);

    /// <summary>
    /// 删除所有指定类型的可编辑区域
    /// </summary>
    /// <param name="editorID">要删除的可编辑区域的编辑器类型</param>
    void DeleteAllEditableRanges(WdEditorType editorID);

    /// <summary>
    /// 删除所有指定标识符的可编辑区域
    /// </summary>
    /// <param name="editorID">要删除的可编辑区域的编辑器标识符，如果为null则表示删除所有编辑器的可编辑区域</param>
    void DeleteAllEditableRanges(string? editorID = null);

    /// <summary>
    /// 根据索引获取范围
    /// </summary>
    /// <param name="start">范围开始索引</param>
    /// <param name="end">范围结束索引</param>
    /// <returns>范围对象</returns>
    IWordRange? Range(int? start = null, int? end = null);

    void RemoveDocumentInformation(WdRemoveDocInfoType removeDocInfoType);

    void LockServerFile();

    void CheckInWithVersion(bool saveChanges = true, string? comments = null, bool makePublic = false, string? versionType = null);

    void SaveAsQuickStyleSet(string fileName);

    void ApplyQuickStyleSet(string name);

    void ApplyQuickStyleSet2(object Style);

    void ApplyDocumentTheme(string fileName);

    void ConvertVietDoc(int codePageOrigin);


    IWordContentControls? SelectContentControlsByTitle(string title);

    IWordContentControls? SelectContentControlsByTag(string tag);

    void ExportAsFixedFormat(string outputFileName, WdExportFormat exportFormat, bool openAfterExport = false,
                            WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
                            WdExportRange range = WdExportRange.wdExportAllDocument, int from = 1, int to = 1,
                            WdExportItem item = WdExportItem.wdExportDocumentContent, bool includeDocProps = false,
                            bool keepIRM = true, WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                            bool docStructureTags = true, bool bitmapMissingFonts = true, bool useISO19005_1 = false,
                            object? fixedFormatExtClassPtr = null);

    void FreezeLayout();

    void UnfreezeLayout();

    void DowngradeDocument();

    void Merge(string fileName, WdMergeTarget? mergeTarget = null, bool? detectFormatChanges = null, WdUseFormattingFrom? useFormattingFrom = null, bool? addToRecentFiles = null);

    bool? CanCheckin();

    void CheckIn(bool saveChanges = true, string? comments = null, bool makePublic = false);


    void Convert();

    void ConvertAutoHyphens();

    void EndReview();

    int? ReturnToLastReadPosition();

    void ReplyWithChanges(bool? showMessage = null);

    void SendForReview(string? recipients = null, string? subject = null, bool? showMessage = null, bool? includeAttachment = null);


    void RemoveLockedStyles();

    void CheckNewSmartTags();

    void RemoveDocumentWorkspaceHeader(string id);

    void RemoveSmartTags();

    void DeleteAllInkAnnotations();

    void RecheckSmartTags();


    void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string id);


    void SetCompatibilityMode([ConvertInt] WdCompatibilityMode Mode);

    IWordXMLNodes? SelectNodes(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    IWordXMLNode? SelectSingleNode(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    void Compare(string name, string? authorName, WdCompareTarget? compareTarget = null,
                bool? detectFormatChanges = null, bool? ignoreAllComparisonWarnings = null, bool? addToRecentFiles = null,
                bool? removePersonalInformation = null, bool? removeDateAndTime = null);

    void TransformDocument(string path, bool dataOnly = true);

    void SendFaxOverInternet(string? recipients = null, string? subject = null, bool? showMessage = null);

    void ResetFormFields();

    void DeleteAllCommentsShown();

    void DeleteAllComments();

    void RejectAllRevisionsShown();

    void AcceptAllRevisionsShown();

    void SetDefaultTableStyle(object Style, bool setInTemplate);

    void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string PasswordEncryptionAlgorithm, int passwordEncryptionKeyLength, bool? PasswordEncryptionFileProperties = null);

    void ReloadAs([ComNamespace("MsCore")] MsoEncoding encoding);

    void WebPagePreview();

    void RemoveTheme();

    void ApplyTheme(string name);

    void DetectLanguage();


    void EditionOptions(WdEditionType Type, WdEditionOption Option, string Name, object Format);

    void MakeCompatibilityDefault();


    void SendFax(string address, string? subject = null);

    void SendMailer(object fileFormat, object priority);

    void CheckConsistency();

    void ClosePrintPreview();

    void PresentIt();

    void UndoClear();

    void ReplyAll();

    void Reply();

    void ForwardMailer();

    void ViewPropertyBrowser();


    bool? Redo(int? times);

    bool? Undo(int? times);

    void ViewCode();

    void AutoFormat();

    object? GetCrossReferenceItems(WdReferenceType referenceType);

    void UpdateSummaryProperties();

    void ToggleFormsDesign();

    void Post();

    void Reload();

    void AddToFavorites();

    void FollowHyperlink(string? address = null, string? subAddress = null, bool? newWindow = null, bool? addHistory = null,
                         string? extraInfo = null, [ComNamespace("MsCore")] MsoExtraInfoMethod? method = null, string? headerInfo = null);

    void CheckSpelling(IWordDictionary? customDictionary = null, bool? ignoreUppercase = null, bool? alwaysSuggest = null,
                    IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null, IWordDictionary? customDictionary4 = null,
                    IWordDictionary? customDictionary5 = null, IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                    IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null, IWordDictionary? customDictionary10 = null);

    void CheckGrammar();

    void UpdateStyles();

    void SetLetterContent(object LetterContent);

    void CopyStylesFromTemplate(string template);

    int? CountNumberedItems(WdNumberType? numberType = null, double? level = null);

    void ConvertNumbersToText(WdNumberType? numberType = null);

    void RemoveNumbers(WdNumberType? numberType = null);

    IWordRange? AutoSummarize(long? length, WdSummaryMode? mode, object? updateProperties);

    IWordRange? GoTo(WdGoToItem? what = null, WdGoToDirection? which = null, int? count = null, string? name = null);
}
