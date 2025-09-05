namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中内置的样式类型枚举
/// </summary>
public enum WdBuiltinStyle
{
    /// <summary>
    /// 正文样式
    /// </summary>
    wdStyleNormal = -1,
    /// <summary>
    /// 信封地址样式
    /// </summary>
    wdStyleEnvelopeAddress = -37,
    /// <summary>
    /// 信封回信地址样式
    /// </summary>
    wdStyleEnvelopeReturn = -38,
    /// <summary>
    /// 正文文本样式
    /// </summary>
    wdStyleBodyText = -67,
    /// <summary>
    /// 标题1样式
    /// </summary>
    wdStyleHeading1 = -2,
    /// <summary>
    /// 标题2样式
    /// </summary>
    wdStyleHeading2 = -3,
    /// <summary>
    /// 标题3样式
    /// </summary>
    wdStyleHeading3 = -4,
    /// <summary>
    /// 标题4样式
    /// </summary>
    wdStyleHeading4 = -5,
    /// <summary>
    /// 标题5样式
    /// </summary>
    wdStyleHeading5 = -6,
    /// <summary>
    /// 标题6样式
    /// </summary>
    wdStyleHeading6 = -7,
    /// <summary>
    /// 标题7样式
    /// </summary>
    wdStyleHeading7 = -8,
    /// <summary>
    /// 标题8样式
    /// </summary>
    wdStyleHeading8 = -9,
    /// <summary>
    /// 标题9样式
    /// </summary>
    wdStyleHeading9 = -10,
    /// <summary>
    /// 索引1样式
    /// </summary>
    wdStyleIndex1 = -11,
    /// <summary>
    /// 索引2样式
    /// </summary>
    wdStyleIndex2 = -12,
    /// <summary>
    /// 索引3样式
    /// </summary>
    wdStyleIndex3 = -13,
    /// <summary>
    /// 索引4样式
    /// </summary>
    wdStyleIndex4 = -14,
    /// <summary>
    /// 索引5样式
    /// </summary>
    wdStyleIndex5 = -15,
    /// <summary>
    /// 索引6样式
    /// </summary>
    wdStyleIndex6 = -16,
    /// <summary>
    /// 索引7样式
    /// </summary>
    wdStyleIndex7 = -17,
    /// <summary>
    /// 索引8样式
    /// </summary>
    wdStyleIndex8 = -18,
    /// <summary>
    /// 索引9样式
    /// </summary>
    wdStyleIndex9 = -19,
    /// <summary>
    /// 目录1样式
    /// </summary>
    wdStyleTOC1 = -20,
    /// <summary>
    /// 目录2样式
    /// </summary>
    wdStyleTOC2 = -21,
    /// <summary>
    /// 目录3样式
    /// </summary>
    wdStyleTOC3 = -22,
    /// <summary>
    /// 目录4样式
    /// </summary>
    wdStyleTOC4 = -23,
    /// <summary>
    /// 目录5样式
    /// </summary>
    wdStyleTOC5 = -24,
    /// <summary>
    /// 目录6样式
    /// </summary>
    wdStyleTOC6 = -25,
    /// <summary>
    /// 目录7样式
    /// </summary>
    wdStyleTOC7 = -26,
    /// <summary>
    /// 目录8样式
    /// </summary>
    wdStyleTOC8 = -27,
    /// <summary>
    /// 目录9样式
    /// </summary>
    wdStyleTOC9 = -28,
    /// <summary>
    /// 正文缩进样式
    /// </summary>
    wdStyleNormalIndent = -29,
    /// <summary>
    /// 脚注文本样式
    /// </summary>
    wdStyleFootnoteText = -30,
    /// <summary>
    /// 批注文本样式
    /// </summary>
    wdStyleCommentText = -31,
    /// <summary>
    /// 页眉样式
    /// </summary>
    wdStyleHeader = -32,
    /// <summary>
    /// 页脚样式
    /// </summary>
    wdStyleFooter = -33,
    /// <summary>
    /// 索引标题样式
    /// </summary>
    wdStyleIndexHeading = -34,
    /// <summary>
    /// 题注样式
    /// </summary>
    wdStyleCaption = -35,
    /// <summary>
    /// 图表目录样式
    /// </summary>
    wdStyleTableOfFigures = -36,
    /// <summary>
    /// 脚注引用样式
    /// </summary>
    wdStyleFootnoteReference = -39,
    /// <summary>
    /// 批注引用样式
    /// </summary>
    wdStyleCommentReference = -40,
    /// <summary>
    /// 行号样式
    /// </summary>
    wdStyleLineNumber = -41,
    /// <summary>
    /// 页码样式
    /// </summary>
    wdStylePageNumber = -42,
    /// <summary>
    /// 尾注引用样式
    /// </summary>
    wdStyleEndnoteReference = -43,
    /// <summary>
    /// 尾注文本样式
    /// </summary>
    wdStyleEndnoteText = -44,
    /// <summary>
    /// 法规索引样式
    /// </summary>
    wdStyleTableOfAuthorities = -45,
    /// <summary>
    /// 宏文本样式
    /// </summary>
    wdStyleMacroText = -46,
    /// <summary>
    /// 法规索引标题样式
    /// </summary>
    wdStyleTOAHeading = -47,
    /// <summary>
    /// 列表样式
    /// </summary>
    wdStyleList = -48,
    /// <summary>
    /// 列表项目符号样式
    /// </summary>
    wdStyleListBullet = -49,
    /// <summary>
    /// 列表编号样式
    /// </summary>
    wdStyleListNumber = -50,
    /// <summary>
    /// 列表2样式
    /// </summary>
    wdStyleList2 = -51,
    /// <summary>
    /// 列表3样式
    /// </summary>
    wdStyleList3 = -52,
    /// <summary>
    /// 列表4样式
    /// </summary>
    wdStyleList4 = -53,
    /// <summary>
    /// 列表5样式
    /// </summary>
    wdStyleList5 = -54,
    /// <summary>
    /// 列表项目符号2样式
    /// </summary>
    wdStyleListBullet2 = -55,
    /// <summary>
    /// 列表项目符号3样式
    /// </summary>
    wdStyleListBullet3 = -56,
    /// <summary>
    /// 列表项目符号4样式
    /// </summary>
    wdStyleListBullet4 = -57,
    /// <summary>
    /// 列表项目符号5样式
    /// </summary>
    wdStyleListBullet5 = -58,
    /// <summary>
    /// 列表编号2样式
    /// </summary>
    wdStyleListNumber2 = -59,
    /// <summary>
    /// 列表编号3样式
    /// </summary>
    wdStyleListNumber3 = -60,
    /// <summary>
    /// 列表编号4样式
    /// </summary>
    wdStyleListNumber4 = -61,
    /// <summary>
    /// 列表编号5样式
    /// </summary>
    wdStyleListNumber5 = -62,
    /// <summary>
    /// 标题样式
    /// </summary>
    wdStyleTitle = -63,
    /// <summary>
    /// 结束语样式
    /// </summary>
    wdStyleClosing = -64,
    /// <summary>
    /// 签名样式
    /// </summary>
    wdStyleSignature = -65,
    /// <summary>
    /// 默认段落字体样式
    /// </summary>
    wdStyleDefaultParagraphFont = -66,
    /// <summary>
    /// 正文文本缩进样式
    /// </summary>
    wdStyleBodyTextIndent = -68,
    /// <summary>
    /// 列表继续样式
    /// </summary>
    wdStyleListContinue = -69,
    /// <summary>
    /// 列表继续2样式
    /// </summary>
    wdStyleListContinue2 = -70,
    /// <summary>
    /// 列表继续3样式
    /// </summary>
    wdStyleListContinue3 = -71,
    /// <summary>
    /// 列表继续4样式
    /// </summary>
    wdStyleListContinue4 = -72,
    /// <summary>
    /// 列表继续5样式
    /// </summary>
    wdStyleListContinue5 = -73,
    /// <summary>
    /// 邮件头样式
    /// </summary>
    wdStyleMessageHeader = -74,
    /// <summary>
    /// 副标题样式
    /// </summary>
    wdStyleSubtitle = -75,
    /// <summary>
    /// 称呼语样式
    /// </summary>
    wdStyleSalutation = -76,
    /// <summary>
    /// 日期样式
    /// </summary>
    wdStyleDate = -77,
    /// <summary>
    /// 正文文本首行缩进样式
    /// </summary>
    wdStyleBodyTextFirstIndent = -78,
    /// <summary>
    /// 正文文本首行缩进2样式
    /// </summary>
    wdStyleBodyTextFirstIndent2 = -79,
    /// <summary>
    /// 注释标题样式
    /// </summary>
    wdStyleNoteHeading = -80,
    /// <summary>
    /// 正文文本2样式
    /// </summary>
    wdStyleBodyText2 = -81,
    /// <summary>
    /// 正文文本3样式
    /// </summary>
    wdStyleBodyText3 = -82,
    /// <summary>
    /// 正文文本缩进2样式
    /// </summary>
    wdStyleBodyTextIndent2 = -83,
    /// <summary>
    /// 正文文本缩进3样式
    /// </summary>
    wdStyleBodyTextIndent3 = -84,
    /// <summary>
    /// 块引用样式
    /// </summary>
    wdStyleBlockQuotation = -85,
    /// <summary>
    /// 超链接样式
    /// </summary>
    wdStyleHyperlink = -86,
    /// <summary>
    /// 已访问超链接样式
    /// </summary>
    wdStyleHyperlinkFollowed = -87,
    /// <summary>
    /// 强调样式
    /// </summary>
    wdStyleStrong = -88,
    /// <summary>
    /// 重点样式
    /// </summary>
    wdStyleEmphasis = -89,
    /// <summary>
    /// 导航窗格样式
    /// </summary>
    wdStyleNavPane = -90,
    /// <summary>
    /// 纯文本样式
    /// </summary>
    wdStylePlainText = -91,
    /// <summary>
    /// HTML正文样式
    /// </summary>
    wdStyleHtmlNormal = -95,
    /// <summary>
    /// HTML首字母缩写样式
    /// </summary>
    wdStyleHtmlAcronym = -96,
    /// <summary>
    /// HTML地址样式
    /// </summary>
    wdStyleHtmlAddress = -97,
    /// <summary>
    /// HTML引用样式
    /// </summary>
    wdStyleHtmlCite = -98,
    /// <summary>
    /// HTML代码样式
    /// </summary>
    wdStyleHtmlCode = -99,
    /// <summary>
    /// HTML定义样式
    /// </summary>
    wdStyleHtmlDfn = -100,
    /// <summary>
    /// HTML键盘样式
    /// </summary>
    wdStyleHtmlKbd = -101,
    /// <summary>
    /// HTML预格式化样式
    /// </summary>
    wdStyleHtmlPre = -102,
    /// <summary>
    /// HTML样本样式
    /// </summary>
    wdStyleHtmlSamp = -103,
    /// <summary>
    /// HTML打字机字体样式
    /// </summary>
    wdStyleHtmlTt = -104,
    /// <summary>
    /// HTML变量样式
    /// </summary>
    wdStyleHtmlVar = -105,
    /// <summary>
    /// 普通表格样式
    /// </summary>
    wdStyleNormalTable = -106,
    /// <summary>
    /// 普通对象样式
    /// </summary>
    wdStyleNormalObject = -158,
    /// <summary>
    /// 表格浅色底纹样式
    /// </summary>
    wdStyleTableLightShading = -159,
    /// <summary>
    /// 表格浅色列表样式
    /// </summary>
    wdStyleTableLightList = -160,
    /// <summary>
    /// 表格浅色网格样式
    /// </summary>
    wdStyleTableLightGrid = -161,
    /// <summary>
    /// 表格中等底纹1样式
    /// </summary>
    wdStyleTableMediumShading1 = -162,
    /// <summary>
    /// 表格中等底纹2样式
    /// </summary>
    wdStyleTableMediumShading2 = -163,
    /// <summary>
    /// 表格中等列表1样式
    /// </summary>
    wdStyleTableMediumList1 = -164,
    /// <summary>
    /// 表格中等列表2样式
    /// </summary>
    wdStyleTableMediumList2 = -165,
    /// <summary>
    /// 表格中等网格1样式
    /// </summary>
    wdStyleTableMediumGrid1 = -166,
    /// <summary>
    /// 表格中等网格2样式
    /// </summary>
    wdStyleTableMediumGrid2 = -167,
    /// <summary>
    /// 表格中等网格3样式
    /// </summary>
    wdStyleTableMediumGrid3 = -168,
    /// <summary>
    /// 表格深色列表样式
    /// </summary>
    wdStyleTableDarkList = -169,
    /// <summary>
    /// 表格彩色底纹样式
    /// </summary>
    wdStyleTableColorfulShading = -170,
    /// <summary>
    /// 表格彩色列表样式
    /// </summary>
    wdStyleTableColorfulList = -171,
    /// <summary>
    /// 表格彩色网格样式
    /// </summary>
    wdStyleTableColorfulGrid = -172,
    /// <summary>
    /// 表格浅色底纹强调1样式
    /// </summary>
    wdStyleTableLightShadingAccent1 = -173,
    /// <summary>
    /// 表格浅色列表强调1样式
    /// </summary>
    wdStyleTableLightListAccent1 = -174,
    /// <summary>
    /// 表格浅色网格强调1样式
    /// </summary>
    wdStyleTableLightGridAccent1 = -175,
    /// <summary>
    /// 表格中等底纹1强调1样式
    /// </summary>
    wdStyleTableMediumShading1Accent1 = -176,
    /// <summary>
    /// 表格中等底纹2强调1样式
    /// </summary>
    wdStyleTableMediumShading2Accent1 = -177,
    /// <summary>
    /// 表格中等列表1强调1样式
    /// </summary>
    wdStyleTableMediumList1Accent1 = -178,
    /// <summary>
    /// 列表段落样式
    /// </summary>
    wdStyleListParagraph = -180,
    /// <summary>
    /// 引用样式
    /// </summary>
    wdStyleQuote = -181,
    /// <summary>
    /// 强烈引用样式
    /// </summary>
    wdStyleIntenseQuote = -182,
    /// <summary>
    /// 淡化强调样式
    /// </summary>
    wdStyleSubtleEmphasis = -261,
    /// <summary>
    /// 强烈强调样式
    /// </summary>
    wdStyleIntenseEmphasis = -262,
    /// <summary>
    /// 淡化参考样式
    /// </summary>
    wdStyleSubtleReference = -263,
    /// <summary>
    /// 强烈参考样式
    /// </summary>
    wdStyleIntenseReference = -264,
    /// <summary>
    /// 书名样式
    /// </summary>
    wdStyleBookTitle = -265,
    /// <summary>
    /// 参考文献样式
    /// </summary>
    wdStyleBibliography = -266,
    /// <summary>
    /// 目录标题样式
    /// </summary>
    wdStyleTocHeading = -267
}