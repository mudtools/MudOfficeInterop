namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word AutoCorrectEntry 对象的封装接口。
/// 用于定义“键入时自动替换”行为，例如将 'teh' 替换为 'the'。
/// </summary>
public interface IWordAutoCorrectEntry : IDisposable
{
    /// <summary>
    /// 获取自动更正条目的名称（即触发词，如 "teh"）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置自动更正条目的替换值（即替换后的内容，如 "the"）。
    /// 注意：设置此属性会修改 Word 中的实际条目。
    /// </summary>
    string Value { get; set; }

    /// <summary>
    /// 删除当前自动更正条目。
    /// </summary>
    void Delete();

    void Apply(IWordRange range);
}