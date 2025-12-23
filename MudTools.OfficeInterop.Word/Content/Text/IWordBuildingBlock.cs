//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word 构建基块（Building Block）对象的封装接口。
/// 提供对名称、内容、类别、类型等属性的访问，并支持删除操作。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordBuildingBlock : IDisposable
{
    /// <summary>
    /// 获取此对象所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父对象（通常是 <see cref="IWordMailMerge"/>）。
    /// </summary>
    object? Parent { get; }

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
    /// 获取或设置构建基块的描述信息。
    /// 描述可用于提供有关构建基块用途或内容的详细信息。
    /// </summary>
    string Description { get; set; }

    /// <summary>
    /// 获取或设置构建基块的唯一标识符。
    /// 此ID用于在Word文档中唯一识别该构建基块。
    /// </summary>
    string ID { get; }

    /// <summary>
    /// 获取或设置插入选项。
    /// 控制构建基块插入到文档时的行为和格式选项。
    /// </summary>
    int InsertOptions { get; set; }

    /// <summary>
    /// 获取构建基块所属的类别（如“常规”、“地址和收件人”等）。
    /// </summary>
    IWordCategory? Category { get; }

    /// <summary>
    /// 获取构建基块的类型（例如“页眉”、“页脚”、“自定义自动图文集”等）。
    /// </summary>
    IWordBuildingBlockType? Type { get; }

    /// <summary>
    /// 删除当前构建基块。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将构建基块插入到指定的范围位置。
    /// </summary>
    /// <param name="where">要插入构建基块的目标范围位置。</param>
    /// <param name="richText">如果为 True，则将构建基块作为 RTF 格式文本插入。 如果为 False，则将构建基块作为纯文本插入。</param>
    /// <returns>插入后的新范围对象，如果插入失败则返回null。</returns>
    IWordRange? Insert(IWordRange where, bool? richText);

}