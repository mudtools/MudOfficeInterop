//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本内容的编码格式，用于Office应用程序中文本的编码和解码操作
/// </summary>
public enum MsoEncoding
{
    /// <summary>泰语编码 (Windows代码页 874)</summary>
    msoEncodingThai = 874,
    /// <summary>日语Shift-JIS编码 (Windows代码页 932)</summary>
    msoEncodingJapaneseShiftJIS = 932,
    /// <summary>简体中文GBK编码 (Windows代码页 936)</summary>
    msoEncodingSimplifiedChineseGBK = 936,
    /// <summary>韩语编码 (Windows代码页 949)</summary>
    msoEncodingKorean = 949,
    /// <summary>繁体中文Big5编码 (Windows代码页 950)</summary>
    msoEncodingTraditionalChineseBig5 = 950,
    /// <summary>Unicode小端序编码</summary>
    msoEncodingUnicodeLittleEndian = 1200,
    /// <summary>Unicode大端序编码</summary>
    msoEncodingUnicodeBigEndian = 1201,
    /// <summary>中欧语言编码 (Windows代码页 1250)</summary>
    msoEncodingCentralEuropean = 1250,
    /// <summary>西里尔编码 (Windows代码页 1251)</summary>
    msoEncodingCyrillic = 1251,
    /// <summary>西欧语言编码 (Windows代码页 1252)</summary>
    msoEncodingWestern = 1252,
    /// <summary>希腊语编码 (Windows代码页 1253)</summary>
    msoEncodingGreek = 1253,
    /// <summary>土耳其语编码 (Windows代码页 1254)</summary>
    msoEncodingTurkish = 1254,
    /// <summary>希伯来语编码 (Windows代码页 1255)</summary>
    msoEncodingHebrew = 1255,
    /// <summary>阿拉伯语编码 (Windows代码页 1256)</summary>
    msoEncodingArabic = 1256,
    /// <summary>波罗的海语言编码 (Windows代码页 1257)</summary>
    msoEncodingBaltic = 1257,
    /// <summary>越南语编码 (Windows代码页 1258)</summary>
    msoEncodingVietnamese = 1258,
    /// <summary>自动检测编码</summary>
    msoEncodingAutoDetect = 50001,
    /// <summary>日语自动检测编码</summary>
    msoEncodingJapaneseAutoDetect = 50932,
    /// <summary>简体中文自动检测编码</summary>
    msoEncodingSimplifiedChineseAutoDetect = 50936,
    /// <summary>韩语自动检测编码</summary>
    msoEncodingKoreanAutoDetect = 50949,
    /// <summary>繁体中文自动检测编码</summary>
    msoEncodingTraditionalChineseAutoDetect = 50950,
    /// <summary>西里尔编码自动检测</summary>
    msoEncodingCyrillicAutoDetect = 51251,
    /// <summary>希腊语自动检测编码</summary>
    msoEncodingGreekAutoDetect = 51253,
    /// <summary>阿拉伯语自动检测编码</summary>
    msoEncodingArabicAutoDetect = 51256,
    /// <summary>ISO 8859-1 拉丁1编码</summary>
    msoEncodingISO88591Latin1 = 28591,
    /// <summary>ISO 8859-2 中欧编码</summary>
    msoEncodingISO88592CentralEurope = 28592,
    /// <summary>ISO 8859-3 拉丁3编码</summary>
    msoEncodingISO88593Latin3 = 28593,
    /// <summary>ISO 8859-4 波罗的海编码</summary>
    msoEncodingISO88594Baltic = 28594,
    /// <summary>ISO 8859-5 西里尔编码</summary>
    msoEncodingISO88595Cyrillic = 28595,
    /// <summary>ISO 8859-6 阿拉伯编码</summary>
    msoEncodingISO88596Arabic = 28596,
    /// <summary>ISO 8859-7 希腊编码</summary>
    msoEncodingISO88597Greek = 28597,
    /// <summary>ISO 8859-8 希伯来编码</summary>
    msoEncodingISO88598Hebrew = 28598,
    /// <summary>ISO 8859-9 土耳其编码</summary>
    msoEncodingISO88599Turkish = 28599,
    /// <summary>ISO 8859-15 拉丁9编码</summary>
    msoEncodingISO885915Latin9 = 28605,
    /// <summary>ISO 8859-8 希伯来逻辑编码</summary>
    msoEncodingISO88598HebrewLogical = 38598,
    /// <summary>ISO 2022 JP 无半角片假名编码</summary>
    msoEncodingISO2022JPNoHalfwidthKatakana = 50220,
    /// <summary>ISO 2022 JP JIS X 0202 1984编码</summary>
    msoEncodingISO2022JPJISX02021984 = 50221,
    /// <summary>ISO 2022 JP JIS X 0201 1989编码</summary>
    msoEncodingISO2022JPJISX02011989 = 50222,
    /// <summary>ISO 2022 KR 韩文编码</summary>
    msoEncodingISO2022KR = 50225,
    /// <summary>ISO 2022 CN 繁体中文编码</summary>
    msoEncodingISO2022CNTraditionalChinese = 50227,
    /// <summary>ISO 2022 CN 简体中文编码</summary>
    msoEncodingISO2022CNSimplifiedChinese = 50229,
    /// <summary>Macintosh 罗马编码</summary>
    msoEncodingMacRoman = 10000,
    /// <summary>Macintosh 日语编码</summary>
    msoEncodingMacJapanese = 10001,
    /// <summary>Macintosh 繁体中文Big5编码</summary>
    msoEncodingMacTraditionalChineseBig5 = 10002,
    /// <summary>Macintosh 韩语编码</summary>
    msoEncodingMacKorean = 10003,
    /// <summary>Macintosh 阿拉伯语编码</summary>
    msoEncodingMacArabic = 10004,
    /// <summary>Macintosh 希伯来语编码</summary>
    msoEncodingMacHebrew = 10005,
    /// <summary>Macintosh 希腊语编码</summary>
    msoEncodingMacGreek1 = 10006,
    /// <summary>Macintosh 西里尔编码</summary>
    msoEncodingMacCyrillic = 10007,
    /// <summary>Macintosh 简体中文GB2312编码</summary>
    msoEncodingMacSimplifiedChineseGB2312 = 10008,
    /// <summary>Macintosh 罗马尼亚语编码</summary>
    msoEncodingMacRomania = 10010,
    /// <summary>Macintosh 乌克兰语编码</summary>
    msoEncodingMacUkraine = 10017,
    /// <summary>Macintosh 拉丁2编码</summary>
    msoEncodingMacLatin2 = 10029,
    /// <summary>Macintosh 冰岛语编码</summary>
    msoEncodingMacIcelandic = 10079,
    /// <summary>Macintosh 土耳其语编码</summary>
    msoEncodingMacTurkish = 10081,
    /// <summary>Macintosh 克罗地亚语编码</summary>
    msoEncodingMacCroatia = 10082,
    /// <summary>EBCDIC 美国/加拿大编码</summary>
    msoEncodingEBCDICUSCanada = 37,
    /// <summary>EBCDIC 国际编码</summary>
    msoEncodingEBCDICInternational = 500,
    /// <summary>EBCDIC 多语言ROECE拉丁2编码</summary>
    msoEncodingEBCDICMultilingualROECELatin2 = 870,
    /// <summary>EBCDIC 现代希腊语编码</summary>
    msoEncodingEBCDICGreekModern = 875,
    /// <summary>EBCDIC 土耳其拉丁5编码</summary>
    msoEncodingEBCDICTurkishLatin5 = 1026,
    /// <summary>EBCDIC 德语编码</summary>
    msoEncodingEBCDICGermany = 20273,
    /// <summary>EBCDIC 丹麦/挪威语编码</summary>
    msoEncodingEBCDICDenmarkNorway = 20277,
    /// <summary>EBCDIC 芬兰/瑞典语编码</summary>
    msoEncodingEBCDICFinlandSweden = 20278,
    /// <summary>EBCDIC 意大利语编码</summary>
    msoEncodingEBCDICItaly = 20280,
    /// <summary>EBCDIC 拉丁美洲/西班牙语编码</summary>
    msoEncodingEBCDICLatinAmericaSpain = 20284,
    /// <summary>EBCDIC 英语编码</summary>
    msoEncodingEBCDICUnitedKingdom = 20285,
    /// <summary>EBCDIC 日语片假名扩展编码</summary>
    msoEncodingEBCDICJapaneseKatakanaExtended = 20290,
    /// <summary>EBCDIC 法语编码</summary>
    msoEncodingEBCDICFrance = 20297,
    /// <summary>EBCDIC 阿拉伯语编码</summary>
    msoEncodingEBCDICArabic = 20420,
    /// <summary>EBCDIC 希腊语编码</summary>
    msoEncodingEBCDICGreek = 20423,
    /// <summary>EBCDIC 希伯来语编码</summary>
    msoEncodingEBCDICHebrew = 20424,
    /// <summary>EBCDIC 韩语扩展编码</summary>
    msoEncodingEBCDICKoreanExtended = 20833,
    /// <summary>EBCDIC 泰语编码</summary>
    msoEncodingEBCDICThai = 20838,
    /// <summary>EBCDIC 冰岛语编码</summary>
    msoEncodingEBCDICIcelandic = 20871,
    /// <summary>EBCDIC 土耳其语编码</summary>
    msoEncodingEBCDICTurkish = 20905,
    /// <summary>EBCDIC 俄语编码</summary>
    msoEncodingEBCDICRussian = 20880,
    /// <summary>EBCDIC 塞尔维亚语/保加利亚语编码</summary>
    msoEncodingEBCDICSerbianBulgarian = 21025,
    /// <summary>EBCDIC 日语片假名扩展和日语编码</summary>
    msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930,
    /// <summary>EBCDIC 美国/加拿大和日语编码</summary>
    msoEncodingEBCDICUSCanadaAndJapanese = 50931,
    /// <summary>EBCDIC 韩语扩展和韩语编码</summary>
    msoEncodingEBCDICKoreanExtendedAndKorean = 50933,
    /// <summary>EBCDIC 简体中文扩展和简体中文编码</summary>
    msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935,
    /// <summary>EBCDIC 美国/加拿大和繁体中文编码</summary>
    msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937,
    /// <summary>EBCDIC 日语拉丁扩展和日语编码</summary>
    msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939,
    /// <summary>OEM 美国编码</summary>
    msoEncodingOEMUnitedStates = 437,
    /// <summary>OEM 希腊437G编码</summary>
    msoEncodingOEMGreek437G = 737,
    /// <summary>OEM 波罗的海编码</summary>
    msoEncodingOEMBaltic = 775,
    /// <summary>OEM 多语言拉丁I编码</summary>
    msoEncodingOEMMultilingualLatinI = 850,
    /// <summary>OEM 多语言拉丁II编码</summary>
    msoEncodingOEMMultilingualLatinII = 852,
    /// <summary>OEM 西里尔编码</summary>
    msoEncodingOEMCyrillic = 855,
    /// <summary>OEM 土耳其语编码</summary>
    msoEncodingOEMTurkish = 857,
    /// <summary>OEM 葡萄牙语编码</summary>
    msoEncodingOEMPortuguese = 860,
    /// <summary>OEM 冰岛语编码</summary>
    msoEncodingOEMIcelandic = 861,
    /// <summary>OEM 希伯来语编码</summary>
    msoEncodingOEMHebrew = 862,
    /// <summary>OEM 加拿大法语编码</summary>
    msoEncodingOEMCanadianFrench = 863,
    /// <summary>OEM 阿拉伯语编码</summary>
    msoEncodingOEMArabic = 864,
    /// <summary>OEM 北欧编码</summary>
    msoEncodingOEMNordic = 865,
    /// <summary>OEM 西里尔II编码</summary>
    msoEncodingOEMCyrillicII = 866,
    /// <summary>OEM 现代希腊语编码</summary>
    msoEncodingOEMModernGreek = 869,
    /// <summary>EUC 日语编码</summary>
    msoEncodingEUCJapanese = 51932,
    /// <summary>EUC 简体中文编码</summary>
    msoEncodingEUCChineseSimplifiedChinese = 51936,
    /// <summary>EUC 韩语编码</summary>
    msoEncodingEUCKorean = 51949,
    /// <summary>EUC 台湾繁体中文编码</summary>
    msoEncodingEUCTaiwaneseTraditionalChinese = 51950,
    /// <summary>ISCII 天城文编码</summary>
    msoEncodingISCIIDevanagari = 57002,
    /// <summary>ISCII 孟加拉文编码</summary>
    msoEncodingISCIIBengali = 57003,
    /// <summary>ISCII 泰米尔文编码</summary>
    msoEncodingISCIITamil = 57004,
    /// <summary>ISCII 泰卢固文编码</summary>
    msoEncodingISCIITelugu = 57005,
    /// <summary>ISCII 阿萨姆文编码</summary>
    msoEncodingISCIIAssamese = 57006,
    /// <summary>ISCII 奥里亚文编码</summary>
    msoEncodingISCIIOriya = 57007,
    /// <summary>ISCII 卡纳达文编码</summary>
    msoEncodingISCIIKannada = 57008,
    /// <summary>ISCII 马拉雅拉姆文编码</summary>
    msoEncodingISCIIMalayalam = 57009,
    /// <summary>ISCII 古吉拉特文编码</summary>
    msoEncodingISCIIGujarati = 57010,
    /// <summary>ISCII 旁遮普文编码</summary>
    msoEncodingISCIIPunjabi = 57011,
    /// <summary>阿拉伯语ASMO编码</summary>
    msoEncodingArabicASMO = 708,
    /// <summary>阿拉伯语透明ASMO编码</summary>
    msoEncodingArabicTransparentASMO = 720,
    /// <summary>韩语Johab编码</summary>
    msoEncodingKoreanJohab = 1361,
    /// <summary>台湾CNS编码</summary>
    msoEncodingTaiwanCNS = 20000,
    /// <summary>台湾TCA编码</summary>
    msoEncodingTaiwanTCA = 20001,
    /// <summary>台湾Eten编码</summary>
    msoEncodingTaiwanEten = 20002,
    /// <summary>台湾IBM5550编码</summary>
    msoEncodingTaiwanIBM5550 = 20003,
    /// <summary>台湾电传编码</summary>
    msoEncodingTaiwanTeleText = 20004,
    /// <summary>台湾Wang编码</summary>
    msoEncodingTaiwanWang = 20005,
    /// <summary>IA5 IRV编码</summary>
    msoEncodingIA5IRV = 20105,
    /// <summary>IA5 德语编码</summary>
    msoEncodingIA5German = 20106,
    /// <summary>IA5 瑞典语编码</summary>
    msoEncodingIA5Swedish = 20107,
    /// <summary>IA5 挪威语编码</summary>
    msoEncodingIA5Norwegian = 20108,
    /// <summary>美国ASCII编码</summary>
    msoEncodingUSASCII = 20127,
    /// <summary>T.61编码</summary>
    msoEncodingT61 = 20261,
    /// <summary>ISO 6937 非间距重音编码</summary>
    msoEncodingISO6937NonSpacingAccent = 20269,
    /// <summary>KOI8-R编码</summary>
    msoEncodingKOI8R = 20866,
    /// <summary>扩展阿尔法小写编码</summary>
    msoEncodingExtAlphaLowercase = 21027,
    /// <summary>KOI8-U编码</summary>
    msoEncodingKOI8U = 21866,
    /// <summary>Europa 3编码</summary>
    msoEncodingEuropa3 = 29001,
    /// <summary>HZ-GB 简体中文编码</summary>
    msoEncodingHZGBSimplifiedChinese = 52936,
    /// <summary>简体中文GB18030编码</summary>
    msoEncodingSimplifiedChineseGB18030 = 54936,
    /// <summary>UTF-7 编码</summary>
    msoEncodingUTF7 = 65000,
    /// <summary>UTF-8 编码</summary>
    msoEncodingUTF8 = 65001
}