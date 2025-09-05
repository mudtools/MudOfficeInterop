namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中字段的类型
/// </summary>
public enum WdFieldType
{
    /// <summary>
    /// 空字段
    /// </summary>
    wdFieldEmpty = -1,
    /// <summary>
    /// 交叉引用字段
    /// </summary>
    wdFieldRef = 3,
    /// <summary>
    /// 索引条目字段
    /// </summary>
    wdFieldIndexEntry = 4,
    /// <summary>
    /// 脚注引用字段
    /// </summary>
    wdFieldFootnoteRef = 5,
    /// <summary>
    /// 集合字段
    /// </summary>
    wdFieldSet = 6,
    /// <summary>
    /// 条件字段
    /// </summary>
    wdFieldIf = 7,
    /// <summary>
    /// 索引字段
    /// </summary>
    wdFieldIndex = 8,
    /// <summary>
    /// 目录条目字段
    /// </summary>
    wdFieldTOCEntry = 9,
    /// <summary>
    /// 样式引用字段
    /// </summary>
    wdFieldStyleRef = 10,
    /// <summary>
    /// 引用文档字段
    /// </summary>
    wdFieldRefDoc = 11,
    /// <summary>
    /// 序列字段
    /// </summary>
    wdFieldSequence = 12,
    /// <summary>
    /// 目录字段
    /// </summary>
    wdFieldTOC = 13,
    /// <summary>
    /// 信息字段
    /// </summary>
    wdFieldInfo = 14,
    /// <summary>
    /// 标题字段
    /// </summary>
    wdFieldTitle = 15,
    /// <summary>
    /// 主题字段
    /// </summary>
    wdFieldSubject = 16,
    /// <summary>
    /// 作者字段
    /// </summary>
    wdFieldAuthor = 17,
    /// <summary>
    /// 关键词字段
    /// </summary>
    wdFieldKeyWord = 18,
    /// <summary>
    /// 注释字段
    /// </summary>
    wdFieldComments = 19,
    /// <summary>
    /// 最后保存者字段
    /// </summary>
    wdFieldLastSavedBy = 20,
    /// <summary>
    /// 创建日期字段
    /// </summary>
    wdFieldCreateDate = 21,
    /// <summary>
    /// 保存日期字段
    /// </summary>
    wdFieldSaveDate = 22,
    /// <summary>
    /// 打印日期字段
    /// </summary>
    wdFieldPrintDate = 23,
    /// <summary>
    /// 修订号字段
    /// </summary>
    wdFieldRevisionNum = 24,
    /// <summary>
    /// 编辑时间字段
    /// </summary>
    wdFieldEditTime = 25,
    /// <summary>
    /// 页数字段
    /// </summary>
    wdFieldNumPages = 26,
    /// <summary>
    /// 字数字段
    /// </summary>
    wdFieldNumWords = 27,
    /// <summary>
    /// 字符数字段
    /// </summary>
    wdFieldNumChars = 28,
    /// <summary>
    /// 文件名字段
    /// </summary>
    wdFieldFileName = 29,
    /// <summary>
    /// 模板字段
    /// </summary>
    wdFieldTemplate = 30,
    /// <summary>
    /// 日期字段
    /// </summary>
    wdFieldDate = 31,
    /// <summary>
    /// 时间字段
    /// </summary>
    wdFieldTime = 32,
    /// <summary>
    /// 页码字段
    /// </summary>
    wdFieldPage = 33,
    /// <summary>
    /// 表达式字段
    /// </summary>
    wdFieldExpression = 34,
    /// <summary>
    /// 引用字段
    /// </summary>
    wdFieldQuote = 35,
    /// <summary>
    /// 包含字段
    /// </summary>
    wdFieldInclude = 36,
    /// <summary>
    /// 页码引用字段
    /// </summary>
    wdFieldPageRef = 37,
    /// <summary>
    /// 询问字段
    /// </summary>
    wdFieldAsk = 38,
    /// <summary>
    /// 填充字段
    /// </summary>
    wdFieldFillIn = 39,
    /// <summary>
    /// 数据字段
    /// </summary>
    wdFieldData = 40,
    /// <summary>
    /// 下一个字段
    /// </summary>
    wdFieldNext = 41,
    /// <summary>
    /// 条件下一个字段
    /// </summary>
    wdFieldNextIf = 42,
    /// <summary>
    /// 条件跳过字段
    /// </summary>
    wdFieldSkipIf = 43,
    /// <summary>
    /// 合并记录字段
    /// </summary>
    wdFieldMergeRec = 44,
    /// <summary>
    /// DDE字段
    /// </summary>
    wdFieldDDE = 45,
    /// <summary>
    /// 自动DDE字段
    /// </summary>
    wdFieldDDEAuto = 46,
    /// <summary>
    /// 术语字段
    /// </summary>
    wdFieldGlossary = 47,
    /// <summary>
    /// 打印字段
    /// </summary>
    wdFieldPrint = 48,
    /// <summary>
    /// 公式字段
    /// </summary>
    wdFieldFormula = 49,
    /// <summary>
    /// 转到按钮字段
    /// </summary>
    wdFieldGoToButton = 50,
    /// <summary>
    /// 宏按钮字段
    /// </summary>
    wdFieldMacroButton = 51,
    /// <summary>
    /// 自动编号大纲字段
    /// </summary>
    wdFieldAutoNumOutline = 52,
    /// <summary>
    /// 自动编号法律字段
    /// </summary>
    wdFieldAutoNumLegal = 53,
    /// <summary>
    /// 自动编号字段
    /// </summary>
    wdFieldAutoNum = 54,
    /// <summary>
    /// 导入字段
    /// </summary>
    wdFieldImport = 55,
    /// <summary>
    /// 链接字段
    /// </summary>
    wdFieldLink = 56,
    /// <summary>
    /// 符号字段
    /// </summary>
    wdFieldSymbol = 57,
    /// <summary>
    /// 嵌入字段
    /// </summary>
    wdFieldEmbed = 58,
    /// <summary>
    /// 合并字段
    /// </summary>
    wdFieldMergeField = 59,
    /// <summary>
    /// 用户名字段
    /// </summary>
    wdFieldUserName = 60,
    /// <summary>
    /// 用户姓名首字母字段
    /// </summary>
    wdFieldUserInitials = 61,
    /// <summary>
    /// 用户地址字段
    /// </summary>
    wdFieldUserAddress = 62,
    /// <summary>
    /// 条形码字段
    /// </summary>
    wdFieldBarCode = 63,
    /// <summary>
    /// 文档变量字段
    /// </summary>
    wdFieldDocVariable = 64,
    /// <summary>
    /// 节字段
    /// </summary>
    wdFieldSection = 65,
    /// <summary>
    /// 节页数字段
    /// </summary>
    wdFieldSectionPages = 66,
    /// <summary>
    /// 包含图片字段
    /// </summary>
    wdFieldIncludePicture = 67,
    /// <summary>
    /// 包含文本字段
    /// </summary>
    wdFieldIncludeText = 68,
    /// <summary>
    /// 文件大小字段
    /// </summary>
    wdFieldFileSize = 69,
    /// <summary>
    /// 文本输入表单字段
    /// </summary>
    wdFieldFormTextInput = 70,
    /// <summary>
    /// 复选框表单字段
    /// </summary>
    wdFieldFormCheckBox = 71,
    /// <summary>
    /// 尾注引用字段
    /// </summary>
    wdFieldNoteRef = 72,
    /// <summary>
    /// 图表目录字段
    /// </summary>
    wdFieldTOA = 73,
    /// <summary>
    /// 图表目录条目字段
    /// </summary>
    wdFieldTOAEntry = 74,
    /// <summary>
    /// 合并序列字段
    /// </summary>
    wdFieldMergeSeq = 75,
    /// <summary>
    /// 私有字段
    /// </summary>
    wdFieldPrivate = 77,
    /// <summary>
    /// 数据库字段
    /// </summary>
    wdFieldDatabase = 78,
    /// <summary>
    /// 自动图文集字段
    /// </summary>
    wdFieldAutoText = 79,
    /// <summary>
    /// 比较字段
    /// </summary>
    wdFieldCompare = 80,
    /// <summary>
    /// 加载项字段
    /// </summary>
    wdFieldAddin = 81,
    /// <summary>
    /// 订阅者字段
    /// </summary>
    wdFieldSubscriber = 82,
    /// <summary>
    /// 下拉表单字段
    /// </summary>
    wdFieldFormDropDown = 83,
    /// <summary>
    /// 高级字段
    /// </summary>
    wdFieldAdvance = 84,
    /// <summary>
    /// 文档属性字段
    /// </summary>
    wdFieldDocProperty = 85,
    /// <summary>
    /// OCX控件字段
    /// </summary>
    wdFieldOCX = 87,
    /// <summary>
    /// 超链接字段
    /// </summary>
    wdFieldHyperlink = 88,
    /// <summary>
    /// 自动图文集列表字段
    /// </summary>
    wdFieldAutoTextList = 89,
    /// <summary>
    /// 列表编号字段
    /// </summary>
    wdFieldListNum = 90,
    /// <summary>
    /// HTML ActiveX字段
    /// </summary>
    wdFieldHTMLActiveX = 91,
    /// <summary>
    /// 双向大纲字段
    /// </summary>
    wdFieldBidiOutline = 92,
    /// <summary>
    /// 地址块字段
    /// </summary>
    wdFieldAddressBlock = 93,
    /// <summary>
    /// 问候语文字段
    /// </summary>
    wdFieldGreetingLine = 94,
    /// <summary>
    /// 形状字段
    /// </summary>
    wdFieldShape = 95,
    /// <summary>
    /// 引用字段
    /// </summary>
    wdFieldCitation = 96,
    /// <summary>
    /// 参考文献字段
    /// </summary>
    wdFieldBibliography = 97,
    /// <summary>
    /// 合并条形码字段
    /// </summary>
    wdFieldMergeBarcode = 98,
    /// <summary>
    /// 显示条形码字段
    /// </summary>
    wdFieldDisplayBarcode = 99
}