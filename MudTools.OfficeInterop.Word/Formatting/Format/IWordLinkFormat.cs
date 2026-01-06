namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.LinkFormat 的接口，用于操作链接格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLinkFormat : IOfficeObject<IWordLinkFormat, MsWord.LinkFormat>, IDisposable
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
    /// 获取或设置一个值，指示在容器文件打开或源文件更改时是否自动更新指定的链接。
    /// </summary>
    bool AutoUpdate { get; set; }

    /// <summary>
    /// 获取指定链接 OLE 对象、图片或字段的源文件名称。
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 获取指定链接 OLE 对象、图片或字段的源文件路径。
    /// </summary>
    string SourcePath { get; }

    /// <summary>
    /// 获取或设置一个值，指示 Field、InlineShape 或 Shape 对象是否已锁定以防止自动更新。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取链接类型。
    /// </summary>
    WdLinkType Type { get; }

    /// <summary>
    /// 获取或设置指定链接 OLE 对象、图片或字段的源文件的完整路径和名称。
    /// </summary>
    string SourceFullName { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示指定的图片是否随文档保存。
    /// </summary>
    bool SavePictureWithDocument { get; set; }

    /// <summary>
    /// 断开源文件与指定 OLE 对象、图片或链接字段之间的链接。
    /// </summary>
    void BreakLink();

    /// <summary>
    /// 更新指定的链接。
    /// </summary>
    void Update();
}