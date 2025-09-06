//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word对话框枚举类型，表示Microsoft Word中可用的各种对话框
/// </summary>
public enum WdWordDialog
{
    /// <summary>
    /// 关于对话框
    /// </summary>
    wdDialogHelpAbout = 9,
    /// <summary>
    /// WordPerfect帮助对话框
    /// </summary>
    wdDialogHelpWordPerfectHelp = 10,
    /// <summary>
    /// 文档统计信息对话框
    /// </summary>
    wdDialogDocumentStatistics = 78,
    /// <summary>
    /// 新建文件对话框
    /// </summary>
    wdDialogFileNew = 79,
    /// <summary>
    /// 打开文件对话框
    /// </summary>
    wdDialogFileOpen = 80,
    /// <summary>
    /// 邮件合并打开数据源对话框
    /// </summary>
    wdDialogMailMergeOpenDataSource = 81,
    /// <summary>
    /// 邮件合并打开页眉数据源对话框
    /// </summary>
    wdDialogMailMergeOpenHeaderSource = 82,
    /// <summary>
    /// 另存为对话框
    /// </summary>
    wdDialogFileSaveAs = 84,
    /// <summary>
    /// 文件摘要信息对话框
    /// </summary>
    wdDialogFileSummaryInfo = 86,
    /// <summary>
    /// 模板对话框
    /// </summary>
    wdDialogToolsTemplates = 87,
    /// <summary>
    /// 打印对话框
    /// </summary>
    wdDialogFilePrint = 88,
    /// <summary>
    /// 打印机设置对话框
    /// </summary>
    wdDialogFilePrintSetup = 97,
    /// <summary>
    /// 查找文件对话框
    /// </summary>
    wdDialogFileFind = 99,
    /// <summary>
    /// 地址字体格式对话框
    /// </summary>
    wdDialogFormatAddrFonts = 103,
    /// <summary>
    /// 选择性粘贴对话框
    /// </summary>
    wdDialogEditPasteSpecial = 111,
    /// <summary>
    /// 查找对话框
    /// </summary>
    wdDialogEditFind = 112,
    /// <summary>
    /// 替换对话框
    /// </summary>
    wdDialogEditReplace = 117,
    /// <summary>
    /// 样式对话框
    /// </summary>
    wdDialogEditStyle = 120,
    /// <summary>
    /// 链接对话框
    /// </summary>
    wdDialogEditLinks = 124,
    /// <summary>
    /// 对象对话框
    /// </summary>
    wdDialogEditObject = 125,
    /// <summary>
    /// 表格转文本对话框
    /// </summary>
    wdDialogTableToText = 128,
    /// <summary>
    /// 文本转表格对话框
    /// </summary>
    wdDialogTextToTable = 127,
    /// <summary>
    /// 插入表格对话框
    /// </summary>
    wdDialogTableInsertTable = 129,
    /// <summary>
    /// 插入单元格对话框
    /// </summary>
    wdDialogTableInsertCells = 130,
    /// <summary>
    /// 插入行对话框
    /// </summary>
    wdDialogTableInsertRow = 131,
    /// <summary>
    /// 删除单元格对话框
    /// </summary>
    wdDialogTableDeleteCells = 133,
    /// <summary>
    /// 拆分单元格对话框
    /// </summary>
    wdDialogTableSplitCells = 137,
    /// <summary>
    /// 行高对话框
    /// </summary>
    wdDialogTableRowHeight = 142,
    /// <summary>
    /// 列宽对话框
    /// </summary>
    wdDialogTableColumnWidth = 143,
    /// <summary>
    /// 自定义对话框
    /// </summary>
    wdDialogToolsCustomize = 152,
    /// <summary>
    /// 插入分隔符对话框
    /// </summary>
    wdDialogInsertBreak = 159,
    /// <summary>
    /// 插入符号对话框
    /// </summary>
    wdDialogInsertSymbol = 162,
    /// <summary>
    /// 插入图片对话框
    /// </summary>
    wdDialogInsertPicture = 163,
    /// <summary>
    /// 插入文件对话框
    /// </summary>
    wdDialogInsertFile = 164,
    /// <summary>
    /// 插入日期时间对话框
    /// </summary>
    wdDialogInsertDateTime = 165,
    /// <summary>
    /// 插入域对话框
    /// </summary>
    wdDialogInsertField = 166,
    /// <summary>
    /// 插入合并域对话框
    /// </summary>
    wdDialogInsertMergeField = 167,
    /// <summary>
    /// 插入书签对话框
    /// </summary>
    wdDialogInsertBookmark = 168,
    /// <summary>
    /// 标记索引项对话框
    /// </summary>
    wdDialogMarkIndexEntry = 169,
    /// <summary>
    /// 插入索引对话框
    /// </summary>
    wdDialogInsertIndex = 170,
    /// <summary>
    /// 插入目录对话框
    /// </summary>
    wdDialogInsertTableOfContents = 171,
    /// <summary>
    /// 插入对象对话框
    /// </summary>
    wdDialogInsertObject = 172,
    /// <summary>
    /// 创建信封对话框
    /// </summary>
    wdDialogToolsCreateEnvelope = 173,
    /// <summary>
    /// 字体格式对话框
    /// </summary>
    wdDialogFormatFont = 174,
    /// <summary>
    /// 段落格式对话框
    /// </summary>
    wdDialogFormatParagraph = 175,
    /// <summary>
    /// 节布局格式对话框
    /// </summary>
    wdDialogFormatSectionLayout = 176,
    /// <summary>
    /// 列格式对话框
    /// </summary>
    wdDialogFormatColumns = 177,
    /// <summary>
    /// 文档布局对话框
    /// </summary>
    wdDialogFileDocumentLayout = 178,
    /// <summary>
    /// 页面设置对话框
    /// </summary>
    wdDialogFilePageSetup = 178,
    /// <summary>
    /// 制表位格式对话框
    /// </summary>
    wdDialogFormatTabs = 179,
    /// <summary>
    /// 样式格式对话框
    /// </summary>
    wdDialogFormatStyle = 180,
    /// <summary>
    /// 定义样式字体对话框
    /// </summary>
    wdDialogFormatDefineStyleFont = 181,
    /// <summary>
    /// 定义样式段落对话框
    /// </summary>
    wdDialogFormatDefineStylePara = 182,
    /// <summary>
    /// 定义样式制表位对话框
    /// </summary>
    wdDialogFormatDefineStyleTabs = 183,
    /// <summary>
    /// 定义样式框架对话框
    /// </summary>
    wdDialogFormatDefineStyleFrame = 184,
    /// <summary>
    /// 定义样式边框对话框
    /// </summary>
    wdDialogFormatDefineStyleBorders = 185,
    /// <summary>
    /// 定义样式语言对话框
    /// </summary>
    wdDialogFormatDefineStyleLang = 186,
    /// <summary>
    /// 图片格式对话框
    /// </summary>
    wdDialogFormatPicture = 187,
    /// <summary>
    /// 语言工具对话框
    /// </summary>
    wdDialogToolsLanguage = 188,
    /// <summary>
    /// 边框和底纹格式对话框
    /// </summary>
    wdDialogFormatBordersAndShading = 189,
    /// <summary>
    /// 框架格式对话框
    /// </summary>
    wdDialogFormatFrame = 190,
    /// <summary>
    /// 同义词库对话框
    /// </summary>
    wdDialogToolsThesaurus = 194,
    /// <summary>
    /// 断字对话框
    /// </summary>
    wdDialogToolsHyphenation = 195,
    /// <summary>
    /// 项目符号和编号对话框
    /// </summary>
    wdDialogToolsBulletsNumbers = 196,
    /// <summary>
    /// 突出显示修订对话框
    /// </summary>
    wdDialogToolsHighlightChanges = 197,
    /// <summary>
    /// 修订对话框
    /// </summary>
    wdDialogToolsRevisions = 197,
    /// <summary>
    /// 比较文档对话框
    /// </summary>
    wdDialogToolsCompareDocuments = 198,
    /// <summary>
    /// 表格排序对话框
    /// </summary>
    wdDialogTableSort = 199,
    /// <summary>
    /// 常规选项对话框
    /// </summary>
    wdDialogToolsOptionsGeneral = 203,
    /// <summary>
    /// 视图选项对话框
    /// </summary>
    wdDialogToolsOptionsView = 204,
    /// <summary>
    /// 高级设置对话框
    /// </summary>
    wdDialogToolsAdvancedSettings = 206,
    /// <summary>
    /// 打印选项对话框
    /// </summary>
    wdDialogToolsOptionsPrint = 208,
    /// <summary>
    /// 保存选项对话框
    /// </summary>
    wdDialogToolsOptionsSave = 209,
    /// <summary>
    /// 拼写和语法选项对话框
    /// </summary>
    wdDialogToolsOptionsSpellingAndGrammar = 211,
    /// <summary>
    /// 用户信息选项对话框
    /// </summary>
    wdDialogToolsOptionsUserInfo = 213,
    /// <summary>
    /// 宏录制对话框
    /// </summary>
    wdDialogToolsMacroRecord = 214,
    /// <summary>
    /// 宏对话框
    /// </summary>
    wdDialogToolsMacro = 215,
    /// <summary>
    /// 激活窗口对话框
    /// </summary>
    wdDialogWindowActivate = 220,
    /// <summary>
    /// 返回地址字体格式对话框
    /// </summary>
    wdDialogFormatRetAddrFonts = 221,
    /// <summary>
    /// 组织器对话框
    /// </summary>
    wdDialogOrganizer = 222,
    /// <summary>
    /// 编辑选项对话框
    /// </summary>
    wdDialogToolsOptionsEdit = 224,
    /// <summary>
    /// 文件位置选项对话框
    /// </summary>
    wdDialogToolsOptionsFileLocations = 225,
    /// <summary>
    /// 字数统计对话框
    /// </summary>
    wdDialogToolsWordCount = 228,
    /// <summary>
    /// 运行控制对话框
    /// </summary>
    wdDialogControlRun = 235,
    /// <summary>
    /// 插入页码对话框
    /// </summary>
    wdDialogInsertPageNumbers = 294,
    /// <summary>
    /// 页码格式对话框
    /// </summary>
    wdDialogFormatPageNumber = 298,
    /// <summary>
    /// 复制文件对话框
    /// </summary>
    wdDialogCopyFile = 300,
    /// <summary>
    /// 更改大小写对话框
    /// </summary>
    wdDialogFormatChangeCase = 322,
    /// <summary>
    /// 更新目录对话框
    /// </summary>
    wdDialogUpdateTOC = 331,
    /// <summary>
    /// 插入数据库对话框
    /// </summary>
    wdDialogInsertDatabase = 341,
    /// <summary>
    /// 表格公式对话框
    /// </summary>
    wdDialogTableFormula = 348,
    /// <summary>
    /// 表单域选项对话框
    /// </summary>
    wdDialogFormFieldOptions = 353,
    /// <summary>
    /// 插入题注对话框
    /// </summary>
    wdDialogInsertCaption = 357,
    /// <summary>
    /// 题注编号对话框
    /// </summary>
    wdDialogInsertCaptionNumbering = 358,
    /// <summary>
    /// 自动题注对话框
    /// </summary>
    wdDialogInsertAutoCaption = 359,
    /// <summary>
    /// 表单域帮助对话框
    /// </summary>
    wdDialogFormFieldHelp = 361,
    /// <summary>
    /// 插入交叉引用对话框
    /// </summary>
    wdDialogInsertCrossReference = 367,
    /// <summary>
    /// 插入脚注对话框
    /// </summary>
    wdDialogInsertFootnote = 370,
    /// <summary>
    /// 脚注选项对话框
    /// </summary>
    wdDialogNoteOptions = 373,
    /// <summary>
    /// 自动更正对话框
    /// </summary>
    wdDialogToolsAutoCorrect = 378,
    /// <summary>
    /// 修订选项对话框
    /// </summary>
    wdDialogToolsOptionsTrackChanges = 386,
    /// <summary>
    /// 转换对象对话框
    /// </summary>
    wdDialogConvertObject = 392,
    /// <summary>
    /// 插入题注对话框
    /// </summary>
    wdDialogInsertAddCaption = 402,
    /// <summary>
    /// 连接对话框
    /// </summary>
    wdDialogConnect = 420,
    /// <summary>
    /// 自定义键盘对话框
    /// </summary>
    wdDialogToolsCustomizeKeyboard = 432,
    /// <summary>
    /// 自定义菜单对话框
    /// </summary>
    wdDialogToolsCustomizeMenus = 433,
    /// <summary>
    /// 合并文档对话框
    /// </summary>
    wdDialogToolsMergeDocuments = 435,
    /// <summary>
    /// 标记目录项对话框
    /// </summary>
    wdDialogMarkTableOfContentsEntry = 442,
    /// <summary>
    /// Mac页面设置GX对话框
    /// </summary>
    wdDialogFileMacPageSetupGX = 444,
    /// <summary>
    /// 打印一份副本对话框
    /// </summary>
    wdDialogFilePrintOneCopy = 445,
    /// <summary>
    /// 编辑框架对话框
    /// </summary>
    wdDialogEditFrame = 458,
    /// <summary>
    /// 标记引文对话框
    /// </summary>
    wdDialogMarkCitation = 463,
    /// <summary>
    /// 目录选项对话框
    /// </summary>
    wdDialogTableOfContentsOptions = 470,
    /// <summary>
    /// 插入引文目录对话框
    /// </summary>
    wdDialogInsertTableOfAuthorities = 471,
    /// <summary>
    /// 插入图表目录对话框
    /// </summary>
    wdDialogInsertTableOfFigures = 472,
    /// <summary>
    /// 插入索引和表格对话框
    /// </summary>
    wdDialogInsertIndexAndTables = 473,
    /// <summary>
    /// 插入表单域对话框
    /// </summary>
    wdDialogInsertFormField = 483,
    /// <summary>
    /// 首字下沉格式对话框
    /// </summary>
    wdDialogFormatDropCap = 488,
    /// <summary>
    /// 创建标签对话框
    /// </summary>
    wdDialogToolsCreateLabels = 489,
    /// <summary>
    /// 保护文档对话框
    /// </summary>
    wdDialogToolsProtectDocument = 503,
    /// <summary>
    /// 样式库格式对话框
    /// </summary>
    wdDialogFormatStyleGallery = 505,
    /// <summary>
    /// 接受/拒绝修订对话框
    /// </summary>
    wdDialogToolsAcceptRejectChanges = 506,
    /// <summary>
    /// WordPerfect帮助选项对话框
    /// </summary>
    wdDialogHelpWordPerfectHelpOptions = 511,
    /// <summary>
    /// 取消文档保护对话框
    /// </summary>
    wdDialogToolsUnprotectDocument = 521,
    /// <summary>
    /// 兼容性选项对话框
    /// </summary>
    wdDialogToolsOptionsCompatibility = 525,
    /// <summary>
    /// 题注选项对话框
    /// </summary>
    wdDialogTableOfCaptionsOptions = 551,
    /// <summary>
    /// 表格自动格式对话框
    /// </summary>
    wdDialogTableAutoFormat = 563,
    /// <summary>
    /// 邮件合并查找记录对话框
    /// </summary>
    wdDialogMailMergeFindRecord = 569,
    /// <summary>
    /// 修订格式对话框
    /// </summary>
    wdDialogReviewAfmtRevisions = 570,
    /// <summary>
    /// 显示比例对话框
    /// </summary>
    wdDialogViewZoom = 577,
    /// <summary>
    /// 保护节对话框
    /// </summary>
    wdDialogToolsProtectSection = 578,
    /// <summary>
    /// 字体替换对话框
    /// </summary>
    wdDialogFontSubstitution = 581,
    /// <summary>
    /// 插入子文档对话框
    /// </summary>
    wdDialogInsertSubdocument = 583,
    /// <summary>
    /// 新建工具栏对话框
    /// </summary>
    wdDialogNewToolbar = 586,
    /// <summary>
    /// 信封和标签对话框
    /// </summary>
    wdDialogToolsEnvelopesAndLabels = 607,
    /// <summary>
    /// 标注格式对话框
    /// </summary>
    wdDialogFormatCallout = 610,
    /// <summary>
    /// 单元格格式对话框
    /// </summary>
    wdDialogTableFormatCell = 612,
    /// <summary>
    /// 自定义菜单栏对话框
    /// </summary>
    wdDialogToolsCustomizeMenuBar = 615,
    /// <summary>
    /// 路由单对话框
    /// </summary>
    wdDialogFileRoutingSlip = 624,
    /// <summary>
    /// 编辑目录类别对话框
    /// </summary>
    wdDialogEditTOACategory = 625,
    /// <summary>
    /// 管理字段对话框
    /// </summary>
    wdDialogToolsManageFields = 631,
    /// <summary>
    /// 对齐到网格对话框
    /// </summary>
    wdDialogDrawSnapToGrid = 633,
    /// <summary>
    /// 绘图对齐对话框
    /// </summary>
    wdDialogDrawAlign = 634,
    /// <summary>
    /// 邮件合并创建数据源对话框
    /// </summary>
    wdDialogMailMergeCreateDataSource = 642,
    /// <summary>
    /// 邮件合并创建页眉源对话框
    /// </summary>
    wdDialogMailMergeCreateHeaderSource = 643,
    /// <summary>
    /// 邮件合并对话框
    /// </summary>
    wdDialogMailMerge = 676,
    /// <summary>
    /// 邮件合并检查对话框
    /// </summary>
    wdDialogMailMergeCheck = 677,
    /// <summary>
    /// 邮件合并助手对话框
    /// </summary>
    wdDialogMailMergeHelper = 680,
    /// <summary>
    /// 邮件合并查询选项对话框
    /// </summary>
    wdDialogMailMergeQueryOptions = 681,
    /// <summary>
    /// Mac页面设置对话框
    /// </summary>
    wdDialogFileMacPageSetup = 685,
    /// <summary>
    /// 列出命令对话框
    /// </summary>
    wdDialogListCommands = 723,
    /// <summary>
    /// 创建发布者对话框
    /// </summary>
    wdDialogEditCreatePublisher = 732,
    /// <summary>
    /// 订阅对话框
    /// </summary>
    wdDialogEditSubscribeTo = 733,
    /// <summary>
    /// 发布选项对话框
    /// </summary>
    wdDialogEditPublishOptions = 735,
    /// <summary>
    /// 订阅选项对话框
    /// </summary>
    wdDialogEditSubscribeOptions = 736,
    /// <summary>
    /// Mac自定义页面设置GX对话框
    /// </summary>
    wdDialogFileMacCustomPageSetupGX = 737,
    /// <summary>
    /// 排版选项对话框
    /// </summary>
    wdDialogToolsOptionsTypography = 739,
    /// <summary>
    /// 自动更正例外对话框
    /// </summary>
    wdDialogToolsAutoCorrectExceptions = 762,
    /// <summary>
    /// 键入时自动套用格式对话框
    /// </summary>
    wdDialogToolsOptionsAutoFormatAsYouType = 778,
    /// <summary>
    /// 邮件合并使用通讯录对话框
    /// </summary>
    wdDialogMailMergeUseAddressBook = 779,
    /// <summary>
    /// 韩文汉字转换对话框
    /// </summary>
    wdDialogToolsHangulHanjaConversion = 784,
    /// <summary>
    /// 模糊选项对话框
    /// </summary>
    wdDialogToolsOptionsFuzzy = 790,
    /// <summary>
    /// 转到旧版对话框
    /// </summary>
    wdDialogEditGoToOld = 811,
    /// <summary>
    /// 插入数字对话框
    /// </summary>
    wdDialogInsertNumber = 812,
    /// <summary>
    /// 信函向导对话框
    /// </summary>
    wdDialogLetterWizard = 821,
    /// <summary>
    /// 项目符号和编号格式对话框
    /// </summary>
    wdDialogFormatBulletsAndNumbering = 824,
    /// <summary>
    /// 拼写和语法检查对话框
    /// </summary>
    wdDialogToolsSpellingAndGrammar = 828,
    /// <summary>
    /// 创建目录对话框
    /// </summary>
    wdDialogToolsCreateDirectory = 833,
    /// <summary>
    /// 表格环绕对话框
    /// </summary>
    wdDialogTableWrapping = 854,
    /// <summary>
    /// 主题格式对话框
    /// </summary>
    wdDialogFormatTheme = 855,
    /// <summary>
    /// 表格属性对话框
    /// </summary>
    wdDialogTableProperties = 861,
    /// <summary>
    /// 电子邮件选项对话框
    /// </summary>
    wdDialogEmailOptions = 863,
    /// <summary>
    /// 创建自动图文集对话框
    /// </summary>
    wdDialogCreateAutoText = 872,
    /// <summary>
    /// 自动摘要对话框
    /// </summary>
    wdDialogToolsAutoSummarize = 874,
    /// <summary>
    /// 语法设置对话框
    /// </summary>
    wdDialogToolsGrammarSettings = 885,
    /// <summary>
    /// 定位对话框
    /// </summary>
    wdDialogEditGoTo = 896,
    /// <summary>
    /// Web选项对话框
    /// </summary>
    wdDialogWebOptions = 898,
    /// <summary>
    /// 插入超链接对话框
    /// </summary>
    wdDialogInsertHyperlink = 925,
    /// <summary>
    /// 自动管理器对话框
    /// </summary>
    wdDialogToolsAutoManager = 915,
    /// <summary>
    /// 文件版本对话框
    /// </summary>
    wdDialogFileVersions = 945,
    /// <summary>
    /// 自动套用格式对话框
    /// </summary>
    wdDialogToolsOptionsAutoFormat = 959,
    /// <summary>
    /// 绘图对象格式对话框
    /// </summary>
    wdDialogFormatDrawingObject = 960,
    /// <summary>
    /// 选项对话框
    /// </summary>
    wdDialogToolsOptions = 974,
    /// <summary>
    /// 适合文字对话框
    /// </summary>
    wdDialogFitText = 983,
    /// <summary>
    /// 编辑自动图文集对话框
    /// </summary>
    wdDialogEditAutoText = 985,
    /// <summary>
    /// 注音指南对话框
    /// </summary>
    wdDialogPhoneticGuide = 986,
    /// <summary>
    /// 字典对话框
    /// </summary>
    wdDialogToolsDictionary = 989,
    /// <summary>
    /// 保存版本对话框
    /// </summary>
    wdDialogFileSaveVersion = 1007,
    /// <summary>
    /// 双向选项对话框
    /// </summary>
    wdDialogToolsOptionsBidi = 1029,
    /// <summary>
    /// 框架集属性对话框
    /// </summary>
    wdDialogFrameSetProperties = 1074,
    /// <summary>
    /// 表格选项对话框
    /// </summary>
    wdDialogTableTableOptions = 1080,
    /// <summary>
    /// 单元格选项对话框
    /// </summary>
    wdDialogTableCellOptions = 1081,
    /// <summary>
    /// IME默认设置对话框
    /// </summary>
    wdDialogIMESetDefault = 1094,
    /// <summary>
    /// 繁简转换对话框
    /// </summary>
    wdDialogTCSCTranslator = 1156,
    /// <summary>
    /// 纵横混排对话框
    /// </summary>
    wdDialogHorizontalInVertical = 1160,
    /// <summary>
    /// 双行合一对话框
    /// </summary>
    wdDialogTwoLinesInOne = 1161,
    /// <summary>
    /// 带圈字符格式对话框
    /// </summary>
    wdDialogFormatEncloseCharacters = 1162,
    /// <summary>
    /// 一致性检查器对话框
    /// </summary>
    wdDialogConsistencyChecker = 1121,
    /// <summary>
    /// 智能标记选项对话框
    /// </summary>
    wdDialogToolsOptionsSmartTag = 1395,
    /// <summary>
    /// 自定义样式对话框
    /// </summary>
    wdDialogFormatStylesCustom = 1248,
    /// <summary>
    /// CSS链接对话框
    /// </summary>
    wdDialogCSSLinks = 1261,
    /// <summary>
    /// 插入Web组件对话框
    /// </summary>
    wdDialogInsertWebComponent = 1324,
    /// <summary>
    /// 编辑复制粘贴选项对话框
    /// </summary>
    wdDialogToolsOptionsEditCopyPaste = 1356,
    /// <summary>
    /// 安全性选项对话框
    /// </summary>
    wdDialogToolsOptionsSecurity = 1361,
    /// <summary>
    /// 搜索对话框
    /// </summary>
    wdDialogSearch = 1363,
    /// <summary>
    /// 显示修复对话框
    /// </summary>
    wdDialogShowRepairs = 1381,
    /// <summary>
    /// 邮件合并插入Ask域对话框
    /// </summary>
    wdDialogMailMergeInsertAsk = 4047,
    /// <summary>
    /// 邮件合并插入FillIn域对话框
    /// </summary>
    wdDialogMailMergeInsertFillIn = 4048,
    /// <summary>
    /// 邮件合并插入If域对话框
    /// </summary>
    wdDialogMailMergeInsertIf = 4049,
    /// <summary>
    /// 邮件合并插入NextIf域对话框
    /// </summary>
    wdDialogMailMergeInsertNextIf = 4053,
    /// <summary>
    /// 邮件合并插入Set域对话框
    /// </summary>
    wdDialogMailMergeInsertSet = 4054,
    /// <summary>
    /// 邮件合并插入SkipIf域对话框
    /// </summary>
    wdDialogMailMergeInsertSkipIf = 4055,
    /// <summary>
    /// 邮件合并域映射对话框
    /// </summary>
    wdDialogMailMergeFieldMapping = 1304,
    /// <summary>
    /// 邮件合并插入地址块对话框
    /// </summary>
    wdDialogMailMergeInsertAddressBlock = 1305,
    /// <summary>
    /// 邮件合并插入问候语对话框
    /// </summary>
    wdDialogMailMergeInsertGreetingLine = 1306,
    /// <summary>
    /// 邮件合并插入域对话框
    /// </summary>
    wdDialogMailMergeInsertFields = 1307,
    /// <summary>
    /// 邮件合并收件人对话框
    /// </summary>
    wdDialogMailMergeRecipients = 1308,
    /// <summary>
    /// 邮件合并查找收件人对话框
    /// </summary>
    wdDialogMailMergeFindRecipient = 1326,
    /// <summary>
    /// 邮件合并设置文档类型对话框
    /// </summary>
    wdDialogMailMergeSetDocumentType = 1339,
    /// <summary>
    /// 标签选项对话框
    /// </summary>
    wdDialogLabelOptions = 1367,
    /// <summary>
    /// XML元素属性对话框
    /// </summary>
    wdDialogXMLElementAttributes = 1460,
    /// <summary>
    /// 架构库对话框
    /// </summary>
    wdDialogSchemaLibrary = 1417,
    /// <summary>
    /// 权限对话框
    /// </summary>
    wdDialogPermission = 1469,
    /// <summary>
    /// 我的权限对话框
    /// </summary>
    wdDialogMyPermission = 1437,
    /// <summary>
    /// XML选项对话框
    /// </summary>
    wdDialogXMLOptions = 1425,
    /// <summary>
    /// 格式限制对话框
    /// </summary>
    wdDialogFormattingRestrictions = 1427,
    /// <summary>
    /// 源管理器对话框
    /// </summary>
    wdDialogSourceManager = 1920,
    /// <summary>
    /// 创建源对话框
    /// </summary>
    wdDialogCreateSource = 1922,
    /// <summary>
    /// 文档检查器对话框
    /// </summary>
    wdDialogDocumentInspector = 1482,
    /// <summary>
    /// 样式管理对话框
    /// </summary>
    wdDialogStyleManagement = 1948,
    /// <summary>
    /// 插入源对话框
    /// </summary>
    wdDialogInsertSource = 2120,
    /// <summary>
    /// 公式识别函数对话框
    /// </summary>
    wdDialogOMathRecognizedFunctions = 2165,
    /// <summary>
    /// 插入占位符对话框
    /// </summary>
    wdDialogInsertPlaceholder = 2348,
    /// <summary>
    /// 构建基块组织器对话框
    /// </summary>
    wdDialogBuildingBlockOrganizer = 2067,
    /// <summary>
    /// 内容控件属性对话框
    /// </summary>
    wdDialogContentControlProperties = 2394,
    /// <summary>
    /// 兼容性检查器对话框
    /// </summary>
    wdDialogCompatibilityChecker = 2439,
    /// <summary>
    /// 导出为固定格式对话框
    /// </summary>
    wdDialogExportAsFixedFormat = 2349,
    /// <summary>
    /// 新建文件2007对话框
    /// </summary>
    wdDialogFileNew2007 = 1116
}