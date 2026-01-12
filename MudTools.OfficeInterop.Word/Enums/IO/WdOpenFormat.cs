//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定打开文档时使用的格式
/// </summary>
public enum WdOpenFormat
{
    /// <summary>
    /// 根据文件扩展名自动确定文件类型
    /// </summary>
    wdOpenFormatAuto = 0,

    /// <summary>
    /// Word文档格式
    /// </summary>
    wdOpenFormatDocument = 1,

    /// <summary>
    /// Word模板格式
    /// </summary>
    wdOpenFormatTemplate = 2,

    /// <summary>
    /// RTF格式（富文本格式）
    /// </summary>
    wdOpenFormatRTF = 3,

    /// <summary>
    /// 文本格式
    /// </summary>
    wdOpenFormatText = 4,

    /// <summary>
    /// Unicode文本格式
    /// </summary>
    wdOpenFormatUnicodeText = 5,

    /// <summary>
    /// 编码文本格式
    /// </summary>
    wdOpenFormatEncodedText = 5,

    /// <summary>
    /// 所有Word文档格式
    /// </summary>
    wdOpenFormatAllWord = 6,

    /// <summary>
    /// 网页格式
    /// </summary>
    wdOpenFormatWebPages = 7,

    /// <summary>
    /// XML格式
    /// </summary>
    wdOpenFormatXML = 8,

    /// <summary>
    /// XML文档格式
    /// </summary>
    wdOpenFormatXMLDocument = 9,

    /// <summary>
    /// 启用宏的XML文档格式
    /// </summary>
    wdOpenFormatXMLDocumentMacroEnabled = 10,

    /// <summary>
    /// XML模板格式
    /// </summary>
    wdOpenFormatXMLTemplate = 11,

    /// <summary>
    /// 启用宏的XML模板格式
    /// </summary>
    wdOpenFormatXMLTemplateMacroEnabled = 12,

    /// <summary>
    /// Word 97文档格式
    /// </summary>
    wdOpenFormatDocument97 = 1,

    /// <summary>
    /// Word 97模板格式
    /// </summary>
    wdOpenFormatTemplate97 = 2,

    /// <summary>
    /// 所有Word模板格式
    /// </summary>
    wdOpenFormatAllWordTemplates = 13,

    /// <summary>
    /// 序列化XML文档格式
    /// </summary>
    wdOpenFormatXMLDocumentSerialized = 14,

    /// <summary>
    /// 启用宏的序列化XML文档格式
    /// </summary>
    wdOpenFormatXMLDocumentMacroEnabledSerialized = 15,

    /// <summary>
    /// 序列化XML模板格式
    /// </summary>
    wdOpenFormatXMLTemplateSerialized = 16,

    /// <summary>
    /// 启用宏的序列化XML模板格式
    /// </summary>
    wdOpenFormatXMLTemplateMacroEnabledSerialized = 17,

    /// <summary>
    /// OpenDocument文本格式
    /// </summary>
    wdOpenFormatOpenDocumentText = 18
}