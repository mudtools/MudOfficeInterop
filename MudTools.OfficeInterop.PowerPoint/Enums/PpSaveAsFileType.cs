//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 保存文件类型枚举
/// </summary>
public enum PpSaveAsFileType
{
    /// <summary>
    /// 演示文稿格式
    /// </summary>
    ppSaveAsPresentation = 1,

    /// <summary>
    /// PowerPoint 7 格式（已过时）
    /// </summary>
    ppSaveAsPowerPoint7 = 2,

    /// <summary>
    /// PowerPoint 4 格式（已过时）
    /// </summary>
    ppSaveAsPowerPoint4 = 3,

    /// <summary>
    /// PowerPoint 3 格式（已过时）
    /// </summary>
    ppSaveAsPowerPoint3 = 4,

    /// <summary>
    /// 模板格式
    /// </summary>
    ppSaveAsTemplate = 5,

    /// <summary>
    /// RTF 格式
    /// </summary>
    ppSaveAsRTF = 6,

    /// <summary>
    /// 放映格式
    /// </summary>
    ppSaveAsShow = 7,

    /// <summary>
    /// 加载项格式
    /// </summary>
    ppSaveAsAddIn = 8,

    /// <summary>
    /// 远东版 PowerPoint 4 格式（已过时）
    /// </summary>
    ppSaveAsPowerPoint4FarEast = 10,

    /// <summary>
    /// 默认格式
    /// </summary>
    ppSaveAsDefault = 11,

    /// <summary>
    /// HTML 格式（已过时）
    /// </summary>
    ppSaveAsHTML = 12,

    /// <summary>
    /// HTML v3 格式（已过时）
    /// </summary>
    ppSaveAsHTMLv3 = 13,

    /// <summary>
    /// HTML 双格式（已过时）
    /// </summary>
    ppSaveAsHTMLDual = 14,

    /// <summary>
    /// 元文件格式
    /// </summary>
    ppSaveAsMetaFile = 15,

    /// <summary>
    /// GIF 格式
    /// </summary>
    ppSaveAsGIF = 16,

    /// <summary>
    /// JPG 格式
    /// </summary>
    ppSaveAsJPG = 17,

    /// <summary>
    /// PNG 格式
    /// </summary>
    ppSaveAsPNG = 18,

    /// <summary>
    /// BMP 格式
    /// </summary>
    ppSaveAsBMP = 19,

    /// <summary>
    /// Web 归档格式（已过时）
    /// </summary>
    ppSaveAsWebArchive = 20,

    /// <summary>
    /// TIF 格式
    /// </summary>
    ppSaveAsTIF = 21,

    /// <summary>
    /// 审阅用演示文稿格式（已过时）
    /// </summary>
    ppSaveAsPresForReview = 22,

    /// <summary>
    /// EMF 格式
    /// </summary>
    ppSaveAsEMF = 23,

    /// <summary>
    /// OpenXML 演示文稿格式
    /// </summary>
    ppSaveAsOpenXMLPresentation = 24,

    /// <summary>
    /// 启用宏的 OpenXML 演示文稿格式
    /// </summary>
    ppSaveAsOpenXMLPresentationMacroEnabled = 25,

    /// <summary>
    /// OpenXML 模板格式
    /// </summary>
    ppSaveAsOpenXMLTemplate = 26,

    /// <summary>
    /// 启用宏的 OpenXML 模板格式
    /// </summary>
    ppSaveAsOpenXMLTemplateMacroEnabled = 27,

    /// <summary>
    /// OpenXML 放映格式
    /// </summary>
    ppSaveAsOpenXMLShow = 28,

    /// <summary>
    /// 启用宏的 OpenXML 放映格式
    /// </summary>
    ppSaveAsOpenXMLShowMacroEnabled = 29,

    /// <summary>
    /// OpenXML 加载项格式
    /// </summary>
    ppSaveAsOpenXMLAddin = 30,

    /// <summary>
    /// OpenXML 主题格式
    /// </summary>
    ppSaveAsOpenXMLTheme = 31,

    /// <summary>
    /// PDF 格式
    /// </summary>
    ppSaveAsPDF = 32,

    /// <summary>
    /// XPS 格式
    /// </summary>
    ppSaveAsXPS = 33,

    /// <summary>
    /// XML 演示文稿格式
    /// </summary>
    ppSaveAsXMLPresentation = 34,

    /// <summary>
    /// OpenDocument 演示文稿格式
    /// </summary>
    ppSaveAsOpenDocumentPresentation = 35,

    /// <summary>
    /// OpenXML 图片演示文稿格式
    /// </summary>
    ppSaveAsOpenXMLPicturePresentation = 36,

    /// <summary>
    /// WMV 格式
    /// </summary>
    ppSaveAsWMV = 37,

    /// <summary>
    /// 严格 OpenXML 演示文稿格式
    /// </summary>
    ppSaveAsStrictOpenXMLPresentation = 38,

    /// <summary>
    /// MP4 格式
    /// </summary>
    ppSaveAsMP4 = 39,

    /// <summary>
    /// 外部转换器格式
    /// </summary>
    ppSaveAsExternalConverter = 64000
}
