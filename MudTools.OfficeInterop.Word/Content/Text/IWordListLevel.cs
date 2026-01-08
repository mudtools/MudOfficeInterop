//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档列表级别接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordListLevel : IOfficeObject<IWordListLevel, MsWord.ListLevel>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取项目在集合中的位置索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置指定列表级别的数字格式。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置指定列表级别数字后面插入的字符。
    /// </summary>
    WdTrailingCharacter TrailingCharacter { get; set; }

    /// <summary>
    /// 获取或设置 ListLevel 对象的数字样式。
    /// </summary>
    WdListNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置指定 ListLevel 对象的数字或项目符号的位置（以磅为单位）。
    /// </summary>
    float NumberPosition { get; set; }

    /// <summary>
    /// 获取或设置表示列表模板中列表级别对齐方式的 WdListLevelAlignment 常量。
    /// </summary>
    WdListLevelAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置指定 ListLevel 对象中换行文本第二行的位置（以磅为单位）。
    /// </summary>
    float TextPosition { get; set; }

    /// <summary>
    /// 获取或设置指定 ListLevel 对象的制表符位置。
    /// </summary>
    float TabPosition { get; set; }

    /// <summary>
    /// 获取或设置指定 ListLevel 对象的起始编号。
    /// </summary>
    int StartAt { get; set; }

    /// <summary>
    /// 获取或设置链接到指定 ListLevel 对象的样式名称。
    /// </summary>
    string LinkedStyle { get; set; }

    /// <summary>
    /// 获取或设置表示指定对象字符格式的 Font 对象。
    /// </summary>
    IWordFont? Font { get; set; }

    /// <summary>
    /// 获取指示创建指定对象的应用程序的 32 位整数。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置必须出现在指定列表级别重新从 1 开始编号之前的列表级别。如果编号每次在列表级别出现时连续进行，则为 false。
    /// </summary>
    int ResetOnHigher { get; set; }

    /// <summary>
    /// 获取表示图片项目符号的 InlineShape 对象。
    /// </summary>
    IWordInlineShape? PictureBullet { get; }

    /// <summary>
    /// 使用图片项目符号格式化段落或段落范围。
    /// </summary>
    /// <param name="fileName">必需。图片文件的路径和名称。</param>
    /// <returns>表示图片项目符号的 InlineShape 对象。</returns>
    IWordInlineShape? ApplyPictureBullet(string fileName);
}