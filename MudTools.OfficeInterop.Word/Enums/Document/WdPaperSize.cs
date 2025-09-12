//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定文档的纸张大小类型，对应Microsoft Word中的纸张尺寸选项
/// </summary>
public enum WdPaperSize
{
    /// <summary>10x14 英寸纸张</summary>
    wdPaper10x14,
    /// <summary>11x17 英寸纸张</summary>
    wdPaper11x17,
    /// <summary>Letter 纸张 (8.5 x 11 英寸)</summary>
    wdPaperLetter,
    /// <summary>Letter Small 纸张 (8.5 x 11 英寸)</summary>
    wdPaperLetterSmall,
    /// <summary>Legal 纸张 (8.5 x 14 英寸)</summary>
    wdPaperLegal,
    /// <summary>Executive 纸张 (7.25 x 10.5 英寸)</summary>
    wdPaperExecutive,
    /// <summary>A3 纸张 (297 x 420 毫米)</summary>
    wdPaperA3,
    /// <summary>A4 纸张 (210 x 297 毫米)</summary>
    wdPaperA4,
    /// <summary>A4 Small 纸张 (210 x 297 毫米)</summary>
    wdPaperA4Small,
    /// <summary>A5 纸张 (148 x 210 毫米)</summary>
    wdPaperA5,
    /// <summary>B4 纸张 (250 x 353 毫米)</summary>
    wdPaperB4,
    /// <summary>B5 纸张 (176 x 250 毫米)</summary>
    wdPaperB5,
    /// <summary>C Sheet 纸张 (17 x 22 英寸)</summary>
    wdPaperCSheet,
    /// <summary>D Sheet 纸张 (22 x 34 英寸)</summary>
    wdPaperDSheet,
    /// <summary>E Sheet 纸张 (34 x 44 英寸)</summary>
    wdPaperESheet,
    /// <summary>German Legal Fanfold 纸张 (8.5 x 13 英寸)</summary>
    wdPaperFanfoldLegalGerman,
    /// <summary>German Standard Fanfold 纸张 (8.5 x 12 英寸)</summary>
    wdPaperFanfoldStdGerman,
    /// <summary>US Standard Fanfold 纸张 (11 x 14.875 英寸)</summary>
    wdPaperFanfoldUS,
    /// <summary>Folio 纸张 (8.5 x 13 英寸)</summary>
    wdPaperFolio,
    /// <summary>Ledger 纸张 (11 x 17 英寸)</summary>
    wdPaperLedger,
    /// <summary>Note 纸张 (8.5 x 11 英寸)</summary>
    wdPaperNote,
    /// <summary>Quarto 纸张 (215 x 275 毫米)</summary>
    wdPaperQuarto,
    /// <summary>Statement 纸张 (5.5 x 8.5 英寸)</summary>
    wdPaperStatement,
    /// <summary>Tabloid 纸张 (11 x 17 英寸)</summary>
    wdPaperTabloid,
    /// <summary>Envelope #9 (3.875 x 8.875 英寸)</summary>
    wdPaperEnvelope9,
    /// <summary>Envelope #10 (4.125 x 9.5 英寸)</summary>
    wdPaperEnvelope10,
    /// <summary>Envelope #11 (4.5 x 10.375 英寸)</summary>
    wdPaperEnvelope11,
    /// <summary>Envelope #12 (4.75 x 11 英寸)</summary>
    wdPaperEnvelope12,
    /// <summary>Envelope #14 (5 x 11.5 英寸)</summary>
    wdPaperEnvelope14,
    /// <summary>Envelope B4 (250 x 353 毫米)</summary>
    wdPaperEnvelopeB4,
    /// <summary>Envelope B5 (176 x 250 毫米)</summary>
    wdPaperEnvelopeB5,
    /// <summary>Envelope B6 (176 x 125 毫米)</summary>
    wdPaperEnvelopeB6,
    /// <summary>Envelope C3 (324 x 458 毫米)</summary>
    wdPaperEnvelopeC3,
    /// <summary>Envelope C4 (229 x 324 毫米)</summary>
    wdPaperEnvelopeC4,
    /// <summary>Envelope C5 (162 x 229 毫米)</summary>
    wdPaperEnvelopeC5,
    /// <summary>Envelope C6 (114 x 162 毫米)</summary>
    wdPaperEnvelopeC6,
    /// <summary>Envelope C65 (114 x 229 毫米)</summary>
    wdPaperEnvelopeC65,
    /// <summary>Envelope DL (110 x 220 毫米)</summary>
    wdPaperEnvelopeDL,
    /// <summary>Italian Envelope (110 x 230 毫米)</summary>
    wdPaperEnvelopeItaly,
    /// <summary>Monarch Envelope (3.875 x 7.5 英寸)</summary>
    wdPaperEnvelopeMonarch,
    /// <summary>Personal Envelope (3.625 x 6.5 英寸)</summary>
    wdPaperEnvelopePersonal,
    /// <summary>自定义纸张大小</summary>
    wdPaperCustom
}