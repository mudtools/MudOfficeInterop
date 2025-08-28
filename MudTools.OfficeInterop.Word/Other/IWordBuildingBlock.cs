namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 构建基块（Building Block）对象的封装接口。
/// 提供对名称、内容、类别、类型等属性的访问，并支持删除操作。
/// </summary>
public interface IWordBuildingBlock : IDisposable
{
    /// <summary>
    /// 获取构建基块的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置构建基块的实际内容文本。
    /// 注意：设置值会替换整个构建基块的内容。
    /// </summary>
    string Value { get; set; }

    /// <summary>
    /// 获取构建基块所属的类别（如“常规”、“地址和收件人”等）。
    /// </summary>
    IWordCategory? Category { get; }

    /// <summary>
    /// 获取构建基块的类型（例如“页眉”、“页脚”、“自定义自动图文集”等）。
    /// </summary>
    string Type { get; }

    /// <summary>
    /// 删除当前构建基块。
    /// </summary>
    void Delete();
}