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

    /// <summary>
    /// 获取文档中的数学公式集合
    /// </summary>
    IWordOMaths? OMaths { get; }

    /// <summary>
    /// 获取或设置二元运算符的换行规则
    /// </summary>
    WdOMathBreakBin OMathBreakBin { get; set; }

    /// <summary>
    /// 获取或设置上标和下标运算符的换行规则
    /// </summary>
    WdOMathBreakSub OMathBreakSub { get; set; }

    /// <summary>
    /// 获取或设置数学公式的对齐方式
    /// </summary>
    WdOMathJc OMathJc { get; set; }

    /// <summary>
    /// 获取或设置数学公式的左边距（单位：磅）
    /// </summary>
    float OMathLeftMargin { get; set; }

    /// <summary>
    /// 获取或设置数学公式的右边距（单位：磅）
    /// </summary>
    float OMathRightMargin { get; set; }

    /// <summary>
    /// 获取或设置数学公式的环绕方式（单位：磅）
    /// </summary>
    float OMathWrap { get; set; }

    /// <summary>
    /// 获取或设置积分符号上标和下标的显示限制
    /// 当为true时，积分符号的上下标以正常方式显示，而不是在符号上方和下方
    /// </summary>
    bool OMathIntSubSupLim { get; set; }

    /// <summary>
    /// 获取或设置N元运算符(如积分、求和等)上标和下标的显示限制
    /// 当为true时，N元运算符的上下标以正常方式显示，而不是在符号上方和下方
    /// </summary>
    bool OMathNarySupSubLim { get; set; }

    /// <summary>
    /// 获取或设置是否使用小分数格式
    /// 当为true时，分数将以较小的格式显示
    /// </summary>
    bool OMathSmallFrac { get; set; }

    /// <summary>
    /// 获取或设置数学公式使用的字体名称
    /// </summary>
    string OMathFontName { get; set; }


    /// <summary>
    /// 获取或设置是否使用数学公式默认设置
    /// </summary>
    bool UseMathDefaults { get; set; }

    /// <summary>
    /// 获取文档中的内容控件集合
    /// </summary>
    IWordContentControls? ContentControls { get; }

    /// <summary>
    /// 获取文档中的参考文献集合
    /// </summary>
    IWordBibliography? Bibliography { get; }


    /// <summary>
    /// 获取文档中的Office脚本集合
    /// </summary>
    IOfficeScripts? Scripts { get; }

    /// <summary>
    /// 获取文档中的XML架构违规节点集合
    /// </summary>
    IWordXMLNodes? XMLSchemaViolations { get; }

    /// <summary>
    /// 获取文档中的XML节点集合
    /// </summary>
    IWordXMLNodes? XMLNodes { get; }

    /// <summary>
    /// 获取XML子节点建议集合
    /// </summary>
    IWordXMLChildNodeSuggestions? ChildNodeSuggestions { get; }


    /// <summary>
    /// 获取文档中的HTML分节集合
    /// </summary>
    IWordHTMLDivisions? HTMLDivisions { get; }

    /// <summary>
    /// 获取文档中的智能标记集合
    /// </summary>
    IWordSmartTags? SmartTags { get; }

    /// <summary>
    /// 获取或设置是否锁定文档主题
    /// </summary>
    bool LockTheme { get; set; }

    /// <summary>
    /// 获取或设置是否锁定快速样式集
    /// </summary>
    bool LockQuickStyleSet { get; set; }

    /// <summary>
    /// 获取原始文档标题
    /// </summary>
    string OriginalDocumentTitle { get; }

    /// <summary>
    /// 获取修订后的文档标题
    /// </summary>
    string RevisedDocumentTitle { get; }

    /// <summary>
    /// 获取或设置格式化显示下一级选项
    /// </summary>
    bool FormattingShowNextLevel { get; set; }

    /// <summary>
    /// 获取或设置格式化是否显示用户样式名称
    /// </summary>
    bool FormattingShowUserStyleName { get; set; }

    /// <summary>
    /// 获取或设置文档是否为最终版本
    /// </summary>
    bool Final { get; set; }

    /// <summary>
    /// 获取文档是否包含VB项目
    /// </summary>
    bool HasVBProject { get; }

    /// <summary>
    /// 获取文档ID
    /// </summary>
    int DocID { get; }

    /// <summary>
    /// 获取当前文档的RSID（修订会话ID）
    /// </summary>
    int CurrentRsid { get; }

    /// <summary>
    /// 获取文档的Word OpenXML内容
    /// </summary>
    string WordOpenXML { get; }

    /// <summary>
    /// 获取文档的兼容性模式
    /// </summary>
    int CompatibilityMode { get; }

    /// <summary>
    /// 获取或设置是否跟踪图表数据点
    /// </summary>
    bool ChartDataPointTrack { get; set; }

    /// <summary>
    /// 获取文档是否正在自动保存中
    /// </summary>
    bool IsInAutosave { get; }

    /// <summary>
    /// 获取或设置是否跟踪格式更改
    /// </summary>
    bool TrackFormatting { get; set; }

    /// <summary>
    /// 获取或设置是否跟踪移动操作
    /// </summary>
    bool TrackMoves { get; set; }

    /// <summary>
    /// 获取或设置样式排序方法
    /// </summary>
    WdStyleSort StyleSortMethod { get; set; }

    /// <summary>
    /// 获取或设置阅读布局的高度
    /// </summary>
    int ReadingLayoutSizeY { get; set; }

    /// <summary>
    /// 获取或设置阅读布局的宽度
    /// </summary>
    int ReadingLayoutSizeX { get; set; }

    /// <summary>
    /// 获取或设置是否移除日期和时间
    /// </summary>
    bool RemoveDateAndTime { get; set; }

    /// <summary>
    /// 获取或设置阅读模式布局是否冻结
    /// </summary>
    bool ReadingModeLayoutFrozen { get; set; }

    /// <summary>
    /// 获取或设置用于保存XML时的XSLT转换路径
    /// </summary>
    string XMLSaveThroughXSLT { get; set; }

    /// <summary>
    /// 获取或设置保存时是否使用XSLT转换
    /// </summary>
    bool XMLUseXSLTWhenSaving { get; set; }

    /// <summary>
    /// 获取或设置是否显示高级XML错误
    /// </summary>
    bool XMLShowAdvancedErrors { get; set; }

    /// <summary>
    /// 获取或设置是否隐藏XML命名空间
    /// </summary>
    bool XMLHideNamespaces { get; set; }

    /// <summary>
    /// 获取或设置是否仅保存XML数据
    /// </summary>
    bool XMLSaveDataOnly { get; set; }

    /// <summary>
    /// 获取或设置是否覆盖自动格式化
    /// </summary>
    bool AutoFormatOverride { get; set; }

    /// <summary>
    /// 获取或设置是否强制使用样式
    /// </summary>
    bool EnforceStyle { get; set; }

    /// <summary>
    /// 获取或设置格式化显示过滤器
    /// </summary>
    WdShowFilter FormattingShowFilter { get; set; }

    /// <summary>
    /// 获取或设置格式化是否显示编号
    /// </summary>
    bool FormattingShowNumbering { get; set; }

    /// <summary>
    /// 获取或设置格式化是否显示段落
    /// </summary>
    bool FormattingShowParagraph { get; set; }

    /// <summary>
    /// 获取或设置格式化是否显示清除
    /// </summary>
    bool FormattingShowClear { get; set; }

    /// <summary>
    /// 获取或设置格式化是否显示字体
    /// </summary>
    bool FormattingShowFont { get; set; }

    /// <summary>
    /// 获取或设置是否嵌入语言数据
    /// </summary>
    bool EmbedLinguisticData { get; set; }

    /// <summary>
    /// 获取或设置是否对文件属性进行密码加密
    /// </summary>
    bool PasswordEncryptionFileProperties { get; }

    /// <summary>
    /// 获取或设置密码加密密钥长度
    /// </summary>
    int PasswordEncryptionKeyLength { get; }

    /// <summary>
    /// 获取或设置密码加密算法
    /// </summary>
    string PasswordEncryptionAlgorithm { get; }

    /// <summary>
    /// 获取或设置密码加密提供程序
    /// </summary>
    string PasswordEncryptionProvider { get; }

    /// <summary>
    /// 获取或设置默认表格样式
    /// </summary>
    object DefaultTableStyle { get; }

    /// <summary>
    /// 获取或设置文本行尾类型
    /// </summary>
    WdLineEndingType TextLineEnding { get; set; }

    /// <summary>
    /// 获取或设置文本编码
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding TextEncoding { get; set; }


    /// <summary>
    /// 获取或设置智能标记是否作为XML属性
    /// </summary>
    bool SmartTagsAsXMLProps { get; set; }

    /// <summary>
    /// 获取或设置是否嵌入智能标记
    /// </summary>
    bool EmbedSmartTags { get; set; }

    /// <summary>
    /// 获取或设置是否移除个人信息
    /// </summary>
    bool RemovePersonalInformation { get; set; }

    /// <summary>
    /// 获取或设置禁用引入的特性
    /// </summary>
    WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter { get; set; }

    /// <summary>
    /// 获取或设置保存编码
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding SaveEncoding { get; set; }

    /// <summary>
    /// 获取或设置打开编码
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding OpenEncoding { get; }

    /// <summary>
    /// 获取或设置默认目标框架
    /// </summary>
    string DefaultTargetFrame { get; set; }

    /// <summary>
    /// 获取或设置是否不嵌入系统字体
    /// </summary>
    bool DoNotEmbedSystemFonts { get; set; }

    /// <summary>
    /// 获取或设置是否禁用特性
    /// </summary>
    bool DisableFeatures { get; set; }

    /// <summary>
    /// 获取或设置VBA是否已签名
    /// </summary>
    bool VBASigned { get; }

    /// <summary>
    /// 获取或设置点击并输入段落样式
    /// </summary>
    [ComPropertyWrap(IsMethod = true)]
    object ClickAndTypeParagraphStyle { get; set; }

    /// <summary>
    /// 获取或设置是否检测语言
    /// </summary>
    bool LanguageDetected { get; set; }

    /// <summary>
    /// 获取活动主题显示名称
    /// </summary>
    string ActiveThemeDisplayName { get; }

    /// <summary>
    /// 获取活动主题
    /// </summary>
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
    /// 获取文档的协同编辑功能接口，用于处理多人同时编辑文档的功能
    /// </summary>
    IWordCoAuthoring? CoAuthoring { get; }

    /// <summary>
    /// 获取文档的自定义XML部件集合接口，用于处理嵌入在文档中的自定义XML数据
    /// </summary>
    IOfficeCustomXMLParts? CustomXMLParts { get; }

    /// <summary>
    /// 获取文档的邮件信封接口，用于处理与邮件发送相关的功能
    /// </summary>
    IOfficeMsoEnvelope? MailEnvelope { get; }

    /// <summary>
    /// 获取文档的HTML项目接口，用于处理HTML格式相关的功能
    /// </summary>
    IOfficeHTMLProject? HTMLProject { get; }

    /// <summary>
    /// 获取文档的智能文档接口，用于处理智能标签和业务智能功能
    /// </summary>
    IOfficeSmartDocument? SmartDocument { get; }

    IOfficeOfficeTheme? DocumentTheme { get; }

    IWordFrameset? Frameset { get; }

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

    /// <summary>
    /// 将文档另存为指定的文件名和格式
    /// </summary>
    /// <param name="fileName">要保存的文件的完整路径和文件名。如果为null，则使用当前文件名</param>
    /// <param name="fileFormat">保存文件的格式，如.docx、.pdf等。如果为null，则使用默认格式</param>
    /// <param name="lockComments">指定是否锁定批注</param>
    /// <param name="password">设置文档的只读密码</param>
    /// <param name="addToRecentFiles">指定是否将文件添加到最近使用的文件列表</param>
    /// <param name="writePassword">设置文档的写入密码</param>
    /// <param name="readOnlyRecommended">指定在打开文档时是否显示只读建议对话框</param>
    /// <param name="embedTrueTypeFonts">指定是否嵌入TrueType字体</param>
    /// <param name="saveNativePictureFormat">指定在保存文档时是否保存图片的原始格式</param>
    /// <param name="saveFormsData">指定是否保存表单数据</param>
    /// <param name="saveAsAOCELetter">指定是否保存为AOCE信件格式</param>
    /// <param name="encoding">指定文件编码方式</param>
    /// <param name="insertLineBreaks">指定是否插入换行符</param>
    /// <param name="allowSubstitutions">指定是否允许替换</param>
    /// <param name="lineEnding">指定行尾类型</param>
    /// <param name="addBiDiMarks">指定是否添加双向语言标记</param>
    void SaveAs(string? fileName = null, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                string? password = null, bool? addToRecentFiles = null, string? writePassword = null,
                bool? readOnlyRecommended = null, bool? embedTrueTypeFonts = null,
                bool? saveNativePictureFormat = null, bool? saveFormsData = null, bool? saveAsAOCELetter = null,
                [ComNamespace("MsCore")] MsoEncoding? encoding = null, bool? insertLineBreaks = null, bool? allowSubstitutions = null,
                WdLineEndingType? lineEnding = null, bool? addBiDiMarks = null);

    /// <summary>
    /// 将文档以指定的文件名和格式保存（扩展版本，包含兼容性模式参数）
    /// </summary>
    /// <param name="fileName">要保存的文件的完整路径和文件名</param>
    /// <param name="fileFormat">保存文件的格式，如.docx、.pdf等。如果为null，则使用默认格式</param>
    /// <param name="lockComments">指定是否锁定批注</param>
    /// <param name="password">设置文档的只读密码</param>
    /// <param name="addToRecentFiles">指定是否将文件添加到最近使用的文件列表</param>
    /// <param name="writePassword">设置文档的写入密码</param>
    /// <param name="readOnlyRecommended">指定在打开文档时是否显示只读建议对话框</param>
    /// <param name="embedTrueTypeFonts">指定是否嵌入TrueType字体</param>
    /// <param name="saveNativePictureFormat">指定在保存文档时是否保存图片的原始格式</param>
    /// <param name="saveFormsData">指定是否保存表单数据</param>
    /// <param name="saveAsAOCELetter">指定是否保存为AOCE信件格式</param>
    /// <param name="encoding">指定文件编码方式</param>
    /// <param name="insertLineBreaks">指定是否插入换行符</param>
    /// <param name="allowSubstitutions">指定是否允许替换</param>
    /// <param name="lineEnding">指定行尾类型</param>
    /// <param name="addBiDiMarks">指定是否添加双向语言标记</param>
    /// <param name="compatibilityMode">指定文档的兼容性模式</param>
    void SaveAs2(string fileName, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                   string? password = null, bool? addToRecentFiles = null, string? writePassword = null,
                   bool? readOnlyRecommended = null, bool? embedTrueTypeFonts = null,
                   bool? saveNativePictureFormat = null, bool? saveFormsData = null,
                   bool? saveAsAOCELetter = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null, bool? insertLineBreaks = null,
                   bool? allowSubstitutions = null, WdLineEndingType? lineEnding = null,
                   bool? addBiDiMarks = null, WdCompatibilityMode? compatibilityMode = null);

    /// <summary>
    /// 将文档副本保存为指定的文件名和格式，不改变当前文档
    /// </summary>
    /// <param name="fileName">要保存的文件副本的完整路径和文件名</param>
    /// <param name="fileFormat">保存文件的格式，如.docx、.pdf等。如果为null，则使用默认格式</param>
    /// <param name="lockComments">指定是否锁定批注</param>
    /// <param name="password">设置文档的只读密码</param>
    /// <param name="addToRecentFiles">指定是否将文件添加到最近使用的文件列表</param>
    /// <param name="writePassword">设置文档的写入密码</param>
    /// <param name="readOnlyRecommended">指定在打开文档时是否显示只读建议对话框</param>
    /// <param name="embedTrueTypeFonts">指定是否嵌入TrueType字体</param>
    /// <param name="saveNativePictureFormat">指定在保存文档时是否保存图片的原始格式</param>
    /// <param name="saveFormsData">指定是否保存表单数据</param>
    /// <param name="saveAsAOCELetter">指定是否保存为AOCE信件格式</param>
    /// <param name="encoding">指定文件编码方式</param>
    /// <param name="insertLineBreaks">指定是否插入换行符</param>
    /// <param name="allowSubstitutions">指定是否允许替换</param>
    /// <param name="lineEnding">指定行尾类型</param>
    /// <param name="addBiDiMarks">指定是否添加双向语言标记</param>
    /// <param name="compatibilityMode">指定文档的兼容性模式</param>
    void SaveCopyAs(string fileName, WdSaveFormat? fileFormat = null, bool? lockComments = null,
                    string? password = null, bool? addToRecentFiles = null, bool? writePassword = null, bool? readOnlyRecommended = null,
                    bool? embedTrueTypeFonts = null, bool? saveNativePictureFormat = null, bool? saveFormsData = null,
                    bool? saveAsAOCELetter = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null,
                    bool? insertLineBreaks = null, bool? allowSubstitutions = null,
                    bool? lineEnding = null, bool? addBiDiMarks = null, WdCompatibilityMode? compatibilityMode = null);



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

    /// <summary>
    /// 移除文档中的特定类型的信息
    /// </summary>
    /// <param name="removeDocInfoType">要移除的文档信息类型</param>
    void RemoveDocumentInformation(WdRemoveDocInfoType removeDocInfoType);

    /// <summary>
    /// 锁定服务器文件
    /// </summary>
    void LockServerFile();

    /// <summary>
    /// 检入文档并指定版本信息
    /// </summary>
    /// <param name="saveChanges">是否保存更改，默认为true</param>
    /// <param name="comments">检入时的注释</param>
    /// <param name="makePublic">是否设为公共版本</param>
    /// <param name="versionType">版本类型</param>
    void CheckInWithVersion(bool saveChanges = true, string? comments = null, bool makePublic = false, string? versionType = null);

    /// <summary>
    /// 将快速样式集保存到外部文件
    /// </summary>
    /// <param name="fileName">要保存的文件名</param>
    void SaveAsQuickStyleSet(string fileName);

    /// <summary>
    /// 应用指定名称的快速样式集
    /// </summary>
    /// <param name="name">要应用的样式集名称</param>
    void ApplyQuickStyleSet(string name);

    /// <summary>
    /// 应用快速样式集的另一种方法
    /// </summary>
    /// <param name="Style">样式对象</param>
    void ApplyQuickStyleSet2(object Style);

    /// <summary>
    /// 应用文档主题
    /// </summary>
    /// <param name="fileName">主题文件路径</param>
    void ApplyDocumentTheme(string fileName);

    /// <summary>
    /// 转换越南文档的编码
    /// </summary>
    /// <param name="codePageOrigin">原始代码页</param>
    void ConvertVietDoc(int codePageOrigin);


    /// <summary>
    /// 根据标题选择内容控件
    /// </summary>
    /// <param name="title">内容控件的标题</param>
    /// <returns>匹配的IWordContentControls对象或null</returns>
    IWordContentControls? SelectContentControlsByTitle(string title);

    /// <summary>
    /// 根据标签选择内容控件
    /// </summary>
    /// <param name="tag">内容控件的标签</param>
    /// <returns>匹配的IWordContentControls对象或null</returns>
    IWordContentControls? SelectContentControlsByTag(string tag);

    /// <summary>
    /// 将文档导出为固定格式（如PDF或XPS）
    /// </summary>
    /// <param name="outputFileName">输出文件名</param>
    /// <param name="exportFormat">导出格式</param>
    /// <param name="openAfterExport">导出后是否打开文件，默认为false</param>
    /// <param name="optimizeFor">优化目标，默认为打印优化</param>
    /// <param name="range">导出范围，默认为整个文档</param>
    /// <param name="from">导出起始页码，默认为1</param>
    /// <param name="to">导出结束页码，默认为1</param>
    /// <param name="item">导出项目，默认为文档内容</param>
    /// <param name="includeDocProps">是否包含文档属性，默认为false</param>
    /// <param name="keepIRM">是否保留IRM信息，默认为true</param>
    /// <param name="createBookmarks">创建书签的方式，默认为不创建</param>
    /// <param name="docStructureTags">是否包含文档结构标签，默认为true</param>
    /// <param name="bitmapMissingFonts">是否用位图替换缺失字体，默认为true</param>
    /// <param name="useISO19005_1">是否使用ISO 19005-1标准（PDF/A），默认为false</param>
    /// <param name="fixedFormatExtClassPtr">固定格式扩展类指针</param>
    void ExportAsFixedFormat(string outputFileName, WdExportFormat exportFormat, bool openAfterExport = false,
                            WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
                            WdExportRange range = WdExportRange.wdExportAllDocument, int from = 1, int to = 1,
                            WdExportItem item = WdExportItem.wdExportDocumentContent, bool includeDocProps = false,
                            bool keepIRM = true, WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                            bool docStructureTags = true, bool bitmapMissingFonts = true, bool useISO19005_1 = false,
                            object? fixedFormatExtClassPtr = null);

    /// <summary>
    /// 冻结文档布局
    /// </summary>
    void FreezeLayout();

    /// <summary>
    /// 解除文档布局冻结
    /// </summary>
    void UnfreezeLayout();

    /// <summary>
    /// 将文档兼容性模式降级
    /// </summary>
    void DowngradeDocument();

    /// <summary>
    /// 合并文档
    /// </summary>
    /// <param name="fileName">要合并的文件名</param>
    /// <param name="mergeTarget">合并目标</param>
    /// <param name="detectFormatChanges">是否检测格式更改</param>
    /// <param name="useFormattingFrom">使用的格式来源</param>
    /// <param name="addToRecentFiles">是否添加到最近文件列表</param>
    void Merge(string fileName, WdMergeTarget? mergeTarget = null, bool? detectFormatChanges = null, WdUseFormattingFrom? useFormattingFrom = null, bool? addToRecentFiles = null);

    /// <summary>
    /// 检查文档是否可以检入
    /// </summary>
    /// <returns>如果可以检入返回true，否则返回false或null</returns>
    bool? CanCheckin();

    /// <summary>
    /// 检入文档
    /// </summary>
    /// <param name="saveChanges">是否保存更改，默认为true</param>
    /// <param name="comments">检入时的注释</param>
    /// <param name="makePublic">是否设为公共版本</param>
    void CheckIn(bool saveChanges = true, string? comments = null, bool makePublic = false);


    /// <summary>
    /// 转换文档格式
    /// </summary>
    void Convert();

    /// <summary>
    /// 转换自动连字符
    /// </summary>
    void ConvertAutoHyphens();

    /// <summary>
    /// 结束文档审阅
    /// </summary>
    void EndReview();

    /// <summary>
    /// 返回到最后阅读位置
    /// </summary>
    /// <returns>返回上次阅读位置，如果失败则返回null</returns>
    int? ReturnToLastReadPosition();

    /// <summary>
    /// 回复修订并包含更改
    /// </summary>
    /// <param name="showMessage">是否显示消息</param>
    void ReplyWithChanges(bool? showMessage = null);

    /// <summary>
    /// 发送文档进行审阅
    /// </summary>
    /// <param name="recipients">收件人列表</param>
    /// <param name="subject">邮件主题</param>
    /// <param name="showMessage">是否显示消息</param>
    /// <param name="includeAttachment">是否包含附件</param>
    void SendForReview(string? recipients = null, string? subject = null, bool? showMessage = null, bool? includeAttachment = null);


    /// <summary>
    /// 移除锁定的样式
    /// </summary>
    void RemoveLockedStyles();

    /// <summary>
    /// 检查新智能标签
    /// </summary>
    void CheckNewSmartTags();

    /// <summary>
    /// 移除文档工作区头部信息
    /// </summary>
    /// <param name="id">要移除的头部ID</param>
    void RemoveDocumentWorkspaceHeader(string id);

    /// <summary>
    /// 移除智能标签
    /// </summary>
    void RemoveSmartTags();

    /// <summary>
    /// 删除所有墨迹注释
    /// </summary>
    void DeleteAllInkAnnotations();

    /// <summary>
    /// 重新检查智能标签
    /// </summary>
    void RecheckSmartTags();


    /// <summary>
    /// 添加文档工作区头部信息
    /// </summary>
    /// <param name="richFormat">是否使用富格式</param>
    /// <param name="url">URL地址</param>
    /// <param name="title">标题</param>
    /// <param name="description">描述</param>
    /// <param name="id">标识符</param>
    void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string id);


    /// <summary>
    /// 设置文档兼容模式
    /// </summary>
    /// <param name="Mode">兼容模式</param>
    void SetCompatibilityMode([ConvertInt] WdCompatibilityMode Mode);

    /// <summary>
    /// 根据XPath选择多个XML节点
    /// </summary>
    /// <param name="XPath">XPath表达式</param>
    /// <param name="prefixMapping">前缀映射</param>
    /// <param name="fastSearchSkippingTextNodes">是否跳过文本节点以快速搜索</param>
    /// <returns>匹配的XML节点集合或null</returns>
    IWordXMLNodes? SelectNodes(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    /// <summary>
    /// 根据XPath选择单个XML节点
    /// </summary>
    /// <param name="XPath">XPath表达式</param>
    /// <param name="prefixMapping">前缀映射</param>
    /// <param name="fastSearchSkippingTextNodes">是否跳过文本节点以快速搜索</param>
    /// <returns>匹配的XML节点或null</returns>
    IWordXMLNode? SelectSingleNode(string XPath, string prefixMapping = "", bool fastSearchSkippingTextNodes = true);

    /// <summary>
    /// 比较文档
    /// </summary>
    /// <param name="name">比较文档的名称</param>
    /// <param name="authorName">作者名称</param>
    /// <param name="compareTarget">比较目标</param>
    /// <param name="detectFormatChanges">是否检测格式更改</param>
    /// <param name="ignoreAllComparisonWarnings">是否忽略所有比较警告</param>
    /// <param name="addToRecentFiles">是否添加到最近文件</param>
    /// <param name="removePersonalInformation">是否移除个人信息</param>
    /// <param name="removeDateAndTime">是否移除日期和时间</param>
    void Compare(string name, string? authorName, WdCompareTarget? compareTarget = null,
                bool? detectFormatChanges = null, bool? ignoreAllComparisonWarnings = null, bool? addToRecentFiles = null,
                bool? removePersonalInformation = null, bool? removeDateAndTime = null);

    /// <summary>
    /// 转换文档使用指定的XSLT路径
    /// </summary>
    /// <param name="path">XSLT路径</param>
    /// <param name="dataOnly">是否仅转换数据</param>
    void TransformDocument(string path, bool dataOnly = true);

    /// <summary>
    /// 通过互联网发送传真
    /// </summary>
    /// <param name="recipients">收件人</param>
    /// <param name="subject">主题</param>
    /// <param name="showMessage">是否显示消息</param>
    void SendFaxOverInternet(string? recipients = null, string? subject = null, bool? showMessage = null);

    /// <summary>
    /// 重置表单字段
    /// </summary>
    void ResetFormFields();

    /// <summary>
    /// 删除所有显示的评论
    /// </summary>
    void DeleteAllCommentsShown();

    /// <summary>
    /// 删除所有评论
    /// </summary>
    void DeleteAllComments();

    /// <summary>
    /// 拒绝所有显示的修订
    /// </summary>
    void RejectAllRevisionsShown();

    /// <summary>
    /// 接受所有显示的修订
    /// </summary>
    void AcceptAllRevisionsShown();

    /// <summary>
    /// 设置默认表格样式
    /// </summary>
    /// <param name="Style">样式对象</param>
    /// <param name="setInTemplate">是否设置在模板中</param>
    void SetDefaultTableStyle(object Style, bool setInTemplate);

    /// <summary>
    /// 设置密码加密选项
    /// </summary>
    /// <param name="passwordEncryptionProvider">密码加密提供程序</param>
    /// <param name="PasswordEncryptionAlgorithm">密码加密算法</param>
    /// <param name="passwordEncryptionKeyLength">密码加密密钥长度</param>
    /// <param name="PasswordEncryptionFileProperties">是否对文件属性进行密码加密</param>
    void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string PasswordEncryptionAlgorithm, int passwordEncryptionKeyLength, bool? PasswordEncryptionFileProperties = null);

    /// <summary>
    /// 重新加载文档使用指定的编码
    /// </summary>
    /// <param name="encoding">编码</param>
    void ReloadAs([ComNamespace("MsCore")] MsoEncoding encoding);

    /// <summary>
    /// 预览网页
    /// </summary>
    void WebPagePreview();

    /// <summary>
    /// 移除主题
    /// </summary>
    void RemoveTheme();

    /// <summary>
    /// 应用主题
    /// </summary>
    /// <param name="name">主题名称</param>
    void ApplyTheme(string name);

    /// <summary>
    /// 检测语言
    /// </summary>
    void DetectLanguage();


    /// <summary>
    /// 设置版式选项
    /// </summary>
    /// <param name="Type">版式类型</param>
    /// <param name="Option">选项</param>
    /// <param name="Name">名称</param>
    /// <param name="Format">格式</param>
    void EditionOptions(WdEditionType Type, WdEditionOption Option, string Name, object Format);

    /// <summary>
    /// 设置默认兼容模式
    /// </summary>
    void MakeCompatibilityDefault();


    /// <summary>
    /// 发送传真
    /// </summary>
    /// <param name="address">地址</param>
    /// <param name="subject">主题</param>
    void SendFax(string address, string? subject = null);

    /// <summary>
    /// 发送邮件
    /// </summary>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="priority">优先级</param>
    void SendMailer(object fileFormat, object priority);

    /// <summary>
    /// 检查一致性
    /// </summary>
    void CheckConsistency();

    /// <summary>
    /// 关闭打印预览
    /// </summary>
    void ClosePrintPreview();

    /// <summary>
    /// 演示
    /// </summary>
    void PresentIt();

    /// <summary>
    /// 清除撤销
    /// </summary>
    void UndoClear();

    /// <summary>
    /// 回复所有
    /// </summary>
    void ReplyAll();

    /// <summary>
    /// 回复
    /// </summary>
    void Reply();

    /// <summary>
    /// 转发邮件
    /// </summary>
    void ForwardMailer();

    /// <summary>
    /// 查看属性浏览器
    /// </summary>
    void ViewPropertyBrowser();


    /// <summary>
    /// 重做
    /// </summary>
    /// <param name="times">重做次数</param>
    /// <returns>是否成功</returns>
    bool? Redo(int? times);

    /// <summary>
    /// 撤销
    /// </summary>
    /// <param name="times">撤销次数</param>
    /// <returns>是否成功</returns>
    bool? Undo(int? times);

    /// <summary>
    /// 查看代码
    /// </summary>
    void ViewCode();

    /// <summary>
    /// 自动格式化
    /// </summary>
    void AutoFormat();

    /// <summary>
    /// 获取交叉引用项
    /// </summary>
    /// <param name="referenceType">引用类型</param>
    /// <returns>交叉引用项</returns>
    object? GetCrossReferenceItems(WdReferenceType referenceType);

    /// <summary>
    /// 更新摘要属性
    /// </summary>
    void UpdateSummaryProperties();

    /// <summary>
    /// 切换表单设计模式
    /// </summary>
    void ToggleFormsDesign();

    /// <summary>
    /// 发布
    /// </summary>
    void Post();

    /// <summary>
    /// 重新加载文档
    /// </summary>
    void Reload();

    /// <summary>
    /// 添加到收藏夹
    /// </summary>
    void AddToFavorites();

    /// <summary>
    /// 跟随超链接
    /// </summary>
    /// <param name="address">地址</param>
    /// <param name="subAddress">子地址</param>
    /// <param name="newWindow">是否在新窗口中打开</param>
    /// <param name="addHistory">是否添加到历史记录</param>
    /// <param name="extraInfo">额外信息</param>
    /// <param name="method">方法</param>
    /// <param name="headerInfo">头部信息</param>
    void FollowHyperlink(string? address = null, string? subAddress = null, bool? newWindow = null, bool? addHistory = null,
                         string? extraInfo = null, [ComNamespace("MsCore")] MsoExtraInfoMethod? method = null, string? headerInfo = null);

    /// <summary>
    /// 检查拼写
    /// </summary>
    /// <param name="customDictionary">自定义词典</param>
    /// <param name="ignoreUppercase">忽略大写字母</param>
    /// <param name="alwaysSuggest">始终建议</param>
    /// <param name="customDictionary2">自定义词典2</param>
    /// <param name="customDictionary3">自定义词典3</param>
    /// <param name="customDictionary4">自定义词典4</param>
    /// <param name="customDictionary5">自定义词典5</param>
    /// <param name="customDictionary6">自定义词典6</param>
    /// <param name="customDictionary7">自定义词典7</param>
    /// <param name="customDictionary8">自定义词典8</param>
    /// <param name="customDictionary9">自定义词典9</param>
    /// <param name="customDictionary10">自定义词典10</param>
    void CheckSpelling(IWordDictionary? customDictionary = null, bool? ignoreUppercase = null, bool? alwaysSuggest = null,
                    IWordDictionary? customDictionary2 = null, IWordDictionary? customDictionary3 = null, IWordDictionary? customDictionary4 = null,
                    IWordDictionary? customDictionary5 = null, IWordDictionary? customDictionary6 = null, IWordDictionary? customDictionary7 = null,
                    IWordDictionary? customDictionary8 = null, IWordDictionary? customDictionary9 = null, IWordDictionary? customDictionary10 = null);

    /// <summary>
    /// 检查语法
    /// </summary>
    void CheckGrammar();

    /// <summary>
    /// 更新样式
    /// </summary>
    void UpdateStyles();

    /// <summary>
    /// 设置信函内容
    /// </summary>
    /// <param name="LetterContent">信函内容对象</param>
    void SetLetterContent(object LetterContent);

    /// <summary>
    /// 从模板复制样式
    /// </summary>
    /// <param name="template">模板路径</param>
    void CopyStylesFromTemplate(string template);

    /// <summary>
    /// 计算编号项目数量
    /// </summary>
    /// <param name="numberType">编号类型</param>
    /// <param name="level">级别</param>
    /// <returns>编号项目数量</returns>
    int? CountNumberedItems(WdNumberType? numberType = null, double? level = null);

    /// <summary>
    /// 将编号转换为文本
    /// </summary>
    /// <param name="numberType">编号类型</param>
    void ConvertNumbersToText(WdNumberType? numberType = null);

    /// <summary>
    /// 移除编号
    /// </summary>
    /// <param name="numberType">编号类型</param>
    void RemoveNumbers(WdNumberType? numberType = null);

    /// <summary>
    /// 自动生成摘要
    /// </summary>
    /// <param name="length">摘要长度</param>
    /// <param name="mode">摘要模式</param>
    /// <param name="updateProperties">更新属性</param>
    /// <returns>摘要范围</returns>
    IWordRange? AutoSummarize(long? length, WdSummaryMode? mode, object? updateProperties);

    /// <summary>
    /// 获取与当前文档关联的信函内容对象，该对象包含通过"信函向导"创建的信函的所有属性和设置。
    /// </summary>
    /// <returns>表示信函内容的对象</returns>
    IWordLetterContent GetLetterContent();

    /// <summary>
    /// 运行Word的信函向导，根据指定的信函内容和向导模式创建或修改信函。
    /// </summary>
    /// <param name="letterContent">信函内容对象，包含信函的各种属性和设置；如果为null，则使用默认设置</param>
    /// <param name="wizardMode">指定是否以交互式向导模式运行；如果为null，则使用默认模式</param>
    void RunLetterWizard(IWordLetterContent? letterContent, bool? wizardMode);

    /// <summary>
    /// 创建一个新的信函内容对象，用于配置通过"信函向导"生成的信函的各个部分和格式。
    /// </summary>
    /// <param name="dateFormat">信函日期的格式</param>
    /// <param name="includeHeaderFooter">是否在信函中包含页眉和页脚</param>
    /// <param name="pageDesign">页面设计模板的名称或路径</param>
    /// <param name="letterStyle">信函的布局样式，如块状布局等</param>
    /// <param name="letterhead">是否在信函中预留预印信头空间</param>
    /// <param name="letterheadLocation">预印信头在信函中的位置</param>
    /// <param name="letterheadSize">预印信头预留的空间大小（以磅为单位）</param>
    /// <param name="recipientName">收件人姓名</param>
    /// <param name="recipientAddress">收件人地址</param>
    /// <param name="salutation">信函的称呼文本</param>
    /// <param name="salutationType">称呼的类型（如正式、非正式等）</param>
    /// <param name="recipientReference">收件人参考行（如"回复："）</param>
    /// <param name="mailingInstructions">邮寄指示文本（如"挂号信"）</param>
    /// <param name="attentionLine">注意行文本</param>
    /// <param name="subject">信函主题</param>
    /// <param name="ccList">抄送（CC）收件人列表</param>
    /// <param name="returnAddress">寄信人地址</param>
    /// <param name="senderName">寄信人姓名</param>
    /// <param name="closing">信函结尾文本（如"此致敬礼"）</param>
    /// <param name="senderCompany">寄信人公司名称</param>
    /// <param name="senderJobTitle">寄信人职位</param>
    /// <param name="senderInitials">寄信人姓名缩写</param>
    /// <param name="enclosureNumber">附件数量</param>
    /// <param name="infoBlock">信息块内容（可选）</param>
    /// <param name="recipientCode">收件人代码（可选）</param>
    /// <param name="recipientGender">收件人性别（可选）</param>
    /// <param name="returnAddressShortForm">简写寄信人地址（可选）</param>
    /// <param name="senderCity">寄信人城市（可选）</param>
    /// <param name="senderCode">寄信人代码（可选）</param>
    /// <param name="senderGender">寄信人性别（可选）</param>
    /// <param name="senderReference">寄信人参考（可选）</param>
    /// <returns>创建的信函内容对象或null</returns>
    IWordLetterContent? CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign,
                         WdLetterStyle letterStyle, bool letterhead, WdLetterheadLocation letterheadLocation,
                         float letterheadSize, string recipientName, string recipientAddress,
                         string salutation, WdSalutationType salutationType, string recipientReference,
                         string mailingInstructions, string attentionLine, string subject, string ccList,
                         string returnAddress, string senderName, string closing, string senderCompany,
                         string senderJobTitle, string senderInitials, int enclosureNumber,
                         string? infoBlock = null, string? recipientCode = null, string? recipientGender = null,
                         string? returnAddressShortForm = null, string? senderCity = null, string? senderCode = null,
                         string? senderGender = null, string? senderReference = null);

    /// <summary>
    /// 根据指定的自定义XML节点选择与之关联的内容控件。
    /// </summary>
    /// <param name="node">用于查找关联内容控件的自定义XML节点</param>
    /// <returns>找到的关联内容控件集合或null</returns>
    IWordContentControls? SelectLinkedControls(IOfficeCustomXMLNode node);

    /// <summary>
    /// 选择未链接到任何自定义XML节点的内容控件。
    /// </summary>
    /// <param name="stream">可选的自定义XML部分，用于指定搜索范围</param>
    /// <returns>未链接的内容控件集合或null</returns>
    IWordContentControls? SelectUnlinkedControls(IOfficeCustomXMLPart stream = null);

    /// <summary>
    /// 跳转到指定位置
    /// </summary>
    /// <param name="what">跳转到的项目</param>
    /// <param name="which">跳转方向</param>
    /// <param name="count">跳转次数</param>
    /// <param name="name">名称</param>
    /// <returns>范围对象</returns>
    IWordRange? GoTo(WdGoToItem? what = null, WdGoToDirection? which = null, int? count = null, string? name = null);
}
