//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 纸张尺寸枚举
/// 用于指定打印时使用的纸张尺寸
/// </summary>
public enum XlPaperSize
{
    /// <summary>
    /// 10x14英寸纸张
    /// </summary>
    xlPaper10x14 = 16,
    
    /// <summary>
    /// 11x17英寸纸张
    /// </summary>
    xlPaper11x17 = 17,
    
    /// <summary>
    /// A3纸张 (297mm x 420mm)
    /// </summary>
    xlPaperA3 = 8,
    
    /// <summary>
    /// A4纸张 (210mm x 297mm)
    /// </summary>
    xlPaperA4 = 9,
    
    /// <summary>
    /// 小号A4纸张
    /// </summary>
    xlPaperA4Small = 10,
    
    /// <summary>
    /// A5纸张 (148mm x 210mm)
    /// </summary>
    xlPaperA5 = 11,
    
    /// <summary>
    /// B4纸张 (250mm x 353mm)
    /// </summary>
    xlPaperB4 = 12,
    
    /// <summary>
    /// B5纸张 (176mm x 250mm)
    /// </summary>
    xlPaperB5 = 13,
    
    /// <summary>
    /// C型纸张
    /// </summary>
    xlPaperCsheet = 24,
    
    /// <summary>
    /// D型纸张
    /// </summary>
    xlPaperDsheet = 25,
    
    /// <summary>
    /// 10号信封 (4.125英寸 x 9.5英寸)
    /// </summary>
    xlPaperEnvelope10 = 20,
    
    /// <summary>
    /// 11号信封
    /// </summary>
    xlPaperEnvelope11 = 21,
    
    /// <summary>
    /// 12号信封
    /// </summary>
    xlPaperEnvelope12 = 22,
    
    /// <summary>
    /// 14号信封
    /// </summary>
    xlPaperEnvelope14 = 23,
    
    /// <summary>
    /// 9号信封 (3.875英寸 x 8.875英寸)
    /// </summary>
    xlPaperEnvelope9 = 19,
    
    /// <summary>
    /// B4信封
    /// </summary>
    xlPaperEnvelopeB4 = 33,
    
    /// <summary>
    /// B5信封
    /// </summary>
    xlPaperEnvelopeB5 = 34,
    
    /// <summary>
    /// B6信封
    /// </summary>
    xlPaperEnvelopeB6 = 35,
    
    /// <summary>
    /// C3信封 (324mm x 458mm)
    /// </summary>
    xlPaperEnvelopeC3 = 29,
    
    /// <summary>
    /// C4信封 (229mm x 324mm)
    /// </summary>
    xlPaperEnvelopeC4 = 30,
    
    /// <summary>
    /// C5信封 (162mm x 229mm)
    /// </summary>
    xlPaperEnvelopeC5 = 28,
    
    /// <summary>
    /// C6信封 (114mm x 162mm)
    /// </summary>
    xlPaperEnvelopeC6 = 31,
    
    /// <summary>
    /// C65信封
    /// </summary>
    xlPaperEnvelopeC65 = 32,
    
    /// <summary>
    /// DL信封 (110mm x 220mm)
    /// </summary>
    xlPaperEnvelopeDL = 27,
    
    /// <summary>
    /// 意大利信封
    /// </summary>
    xlPaperEnvelopeItaly = 36,
    
    /// <summary>
    /// Monarch信封 (3.875英寸 x 7.5英寸)
    /// </summary>
    xlPaperEnvelopeMonarch = 37,
    
    /// <summary>
    /// 个人信封 (3.625英寸 x 6.5英寸)
    /// </summary>
    xlPaperEnvelopePersonal = 38,
    
    /// <summary>
    /// E型纸张
    /// </summary>
    xlPaperEsheet = 26,
    
    /// <summary>
    /// Executive纸张 (7.25英寸 x 10.5英寸)
    /// </summary>
    xlPaperExecutive = 7,
    
    /// <summary>
    /// 德国法律标准Fanfold纸张
    /// </summary>
    xlPaperFanfoldLegalGerman = 41,
    
    /// <summary>
    /// 德国标准Fanfold纸张
    /// </summary>
    xlPaperFanfoldStdGerman = 40,
    
    /// <summary>
    /// 美国标准Fanfold纸张 (14.875英寸 x 11英寸)
    /// </summary>
    xlPaperFanfoldUS = 39,
    
    /// <summary>
    /// Folio纸张 (8.5英寸 x 13英寸)
    /// </summary>
    xlPaperFolio = 14,
    
    /// <summary>
    /// Ledger纸张 (17英寸 x 11英寸)
    /// </summary>
    xlPaperLedger = 4,
    
    /// <summary>
    /// Legal纸张 (8.5英寸 x 14英寸)
    /// </summary>
    xlPaperLegal = 5,
    
    /// <summary>
    /// Letter纸张 (8.5英寸 x 11英寸)
    /// </summary>
    xlPaperLetter = 1,
    
    /// <summary>
    /// 小号Letter纸张
    /// </summary>
    xlPaperLetterSmall = 2,
    
    /// <summary>
    /// Note纸张 (8.5英寸 x 11英寸)
    /// </summary>
    xlPaperNote = 18,
    
    /// <summary>
    /// Quarto纸张 (215mm x 275mm)
    /// </summary>
    xlPaperQuarto = 15,
    
    /// <summary>
    /// Statement纸张 (5.5英寸 x 8.5英寸)
    /// </summary>
    xlPaperStatement = 6,
    
    /// <summary>
    /// Tabloid纸张 (11英寸 x 17英寸)
    /// </summary>
    xlPaperTabloid = 3,
    
    /// <summary>
    /// 用户自定义纸张
    /// </summary>
    xlPaperUser = 256
}