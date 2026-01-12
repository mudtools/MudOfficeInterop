
namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 指定Excel工作簿的文件格式
/// </summary>
public enum XlFileFormat
{
    /// <summary>Excel 97-2003 增加载入项</summary>
    xlAddIn = 18,
    /// <summary>逗号分隔值文件</summary>
    xlCSV = 6,
    /// <summary>Macintosh 逗号分隔值文件</summary>
    xlCSVMac = 22,
    /// <summary>MS-DOS 逗号分隔值文件</summary>
    xlCSVMSDOS = 24,
    /// <summary>Windows 逗号分隔值文件</summary>
    xlCSVWindows = 23,
    /// <summary>dBase II 格式文件</summary>
    xlDBF2 = 7,
    /// <summary>dBase III 格式文件</summary>
    xlDBF3 = 8,
    /// <summary>dBase IV 格式文件</summary>
    xlDBF4 = 11,
    /// <summary>数据交换格式文件</summary>
    xlDIF = 9,
    /// <summary>Excel 2.0 格式文件</summary>
    xlExcel2 = 16,
    /// <summary>远东地区 Excel 2.0 格式文件</summary>
    xlExcel2FarEast = 27,
    /// <summary>Excel 3.0 格式文件</summary>
    xlExcel3 = 29,
    /// <summary>Excel 4.0 格式文件</summary>
    xlExcel4 = 33,
    /// <summary>Excel 5.0 格式文件</summary>
    xlExcel5 = 39,
    /// <summary>Excel 95 格式文件（与 Excel 5.0 相同）</summary>
    xlExcel7 = 39,
    /// <summary>Excel 97-95 格式文件</summary>
    xlExcel9795 = 43,
    /// <summary>Excel 4.0 工作簿格式</summary>
    xlExcel4Workbook = 35,
    /// <summary>国际增加载入项</summary>
    xlIntlAddIn = 26,
    /// <summary>国际宏</summary>
    xlIntlMacro = 25,
    /// <summary>标准 Excel 工作簿格式</summary>
    xlWorkbookNormal = -4143,
    /// <summary>符号链接格式</summary>
    xlSYLK = 2,
    /// <summary>Excel 模板格式</summary>
    xlTemplate = 17,
    /// <summary>当前平台文本格式</summary>
    xlCurrentPlatformText = -4158,
    /// <summary>Macintosh 文本格式</summary>
    xlTextMac = 19,
    /// <summary>MS-DOS 文本格式</summary>
    xlTextMSDOS = 21,
    /// <summary>打印机文本格式</summary>
    xlTextPrinter = 36,
    /// <summary>Windows 文本格式</summary>
    xlTextWindows = 20,
    /// <summary>Lotus 2.x 格式</summary>
    xlWJ2WD1 = 14,
    /// <summary>Lotus 1-2-3 格式</summary>
    xlWK1 = 5,
    /// <summary>Lotus 1-2-3 格式（所有）</summary>
    xlWK1ALL = 31,
    /// <summary>Lotus 1-2-3 格式（FMT）</summary>
    xlWK1FMT = 30,
    /// <summary>Lotus 1-2-3 格式（版本3）</summary>
    xlWK3 = 15,
    /// <summary>Lotus 1-2-3 格式（版本4）</summary>
    xlWK4 = 38,
    /// <summary>Lotus 1-2-3 格式（版本3 FM3）</summary>
    xlWK3FM3 = 32,
    /// <summary>Lotus Symphony 格式</summary>
    xlWKS = 4,
    /// <summary>远东地区 Lotus 工作文件格式</summary>
    xlWorks2FarEast = 28,
    /// <summary>Quattro Pro 格式</summary>
    xlWQ1 = 34,
    /// <summary>Excel 97 格式</summary>
    xlWJ3 = 40,
    /// <summary>Excel 97 FJ3 格式</summary>
    xlWJ3FJ3 = 41,
    /// <summary>Unicode 文本格式</summary>
    xlUnicodeText = 42,
    /// <summary>HTML 格式</summary>
    xlHtml = 44,
    /// <summary>Web 档案格式</summary>
    xlWebArchive = 45,
    /// <summary>XML 电子表格格式</summary>
    xlXMLSpreadsheet = 46,
    /// <summary>Excel 2007 格式</summary>
    xlExcel12 = 50,
    /// <summary>Office Open XML 工作簿格式</summary>
    xlOpenXMLWorkbook = 51,
    /// <summary>启用宏的 Office Open XML 工作簿格式</summary>
    xlOpenXMLWorkbookMacroEnabled = 52,
    /// <summary>启用宏的 Office Open XML 模板格式</summary>
    xlOpenXMLTemplateMacroEnabled = 53,
    /// <summary>Excel 97-2003 模板格式</summary>
    xlTemplate8 = 17,
    /// <summary>Office Open XML 模板格式</summary>
    xlOpenXMLTemplate = 54,
    /// <summary>Excel 97-2003 增加载入项</summary>
    xlAddIn8 = 18,
    /// <summary>Office Open XML 增加载入项</summary>
    xlOpenXMLAddIn = 55,
    /// <summary>Excel 97-2003 格式</summary>
    xlExcel8 = 56,
    /// <summary>开放文档电子表格格式</summary>
    xlOpenDocumentSpreadsheet = 60,
    /// <summary>严格的 Office Open XML 工作簿格式</summary>
    xlOpenXMLStrictWorkbook = 61,
    /// <summary>默认工作簿格式（Office Open XML）</summary>
    xlWorkbookDefault = 51
}