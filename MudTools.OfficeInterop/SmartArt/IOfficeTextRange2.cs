//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeTextRange2 : IEnumerable<IOfficeTextRange2>, IDisposable
{
    object Parent { get; }

    string Text { get; set; }

    int Count { get; }

    IOfficeTextRange2? this[int index] { get; }

    IOfficeTextRange2? this[string name] { get; }

    IOfficeTextRange2 Paragraphs { get; }

    IOfficeTextRange2 Sentences { get; }

    IOfficeTextRange2 Words { get; }

    IOfficeTextRange2 Characters { get; }

    IOfficeTextRange2 Lines { get; }

    IOfficeTextRange2 Runs { get; }

    IOfficeTextRange2 MathZones { get; }

    int Length { get; }

    int Start { get; }

    float BoundLeft { get; }

    float BoundTop { get; }

    float BoundWidth { get; }

    float BoundHeight { get; }

    MsoLanguageID LanguageID { get; set; }

    IOfficeTextRange2 TrimText();

    IOfficeTextRange2 InsertAfter(string newText = "");

    IOfficeTextRange2 InsertBefore(string newText = "");

    IOfficeTextRange2 InsertSymbol(string fontName, int charNumber, [ConvertTriState] bool unicode = false);

    void Select();

    void Cut();

    void Copy();

    void Delete();

    IOfficeTextRange2 Paste();

    IOfficeTextRange2 PasteSpecial(MsoClipboardFormat format);

    void ChangeCase(MsoTextChangeCase Type);

    void AddPeriods();

    void RemovePeriods();

    IOfficeTextRange2 Find(string findWhat, int after = 0, [ConvertTriState] bool MatchCase = false, [ConvertTriState] bool wholeWords = false);

    IOfficeTextRange2 Replace(string findWhat, string replaceWhat, int after = 0, [ConvertTriState] bool matchCase = false, [ConvertTriState] bool wholeWords = false);

    void RtlRun();

    void LtrRun();

    IOfficeTextRange2 InsertChartField(MsoChartFieldType chartFieldType, string formula = "", int position = -1);

}
