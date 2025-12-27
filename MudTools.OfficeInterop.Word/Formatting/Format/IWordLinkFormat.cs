namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.LinkFormat 的接口，用于操作链接格式。
/// </summary>
public interface IWordLinkFormat : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置链接源文件的路径。
    /// </summary>
    string SourceFullName { get; set; }

    /// <summary>
    /// 获取链接源文件的名称（不含路径）。
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 获取链接源文件的应用程序名称。
    /// </summary>
    string SourcePath { get; }

    /// <summary>
    /// 获取或设置是否自动更新链接。
    /// </summary>
    bool? AutoUpdate { get; set; }

    /// <summary>
    /// 获取链接对象的父对象类型。
    /// </summary>
    string ParentType { get; }

    /// <summary>
    /// 获取链接对象的父对象名称。
    /// </summary>
    string ParentName { get; }

    /// <summary>
    /// 获取链接的更新类型。
    /// </summary>
    WdLinkType Type { get; }

    /// <summary>
    /// 更新链接内容。
    /// </summary>
    /// <returns>是否更新成功。</returns>
    bool Update();

    /// <summary>
    /// 断开链接。
    /// </summary>
    /// <returns>是否断开成功。</returns>
    bool BreakLink();

    /// <summary>
    /// 重新链接到新的源文件。
    /// </summary>
    /// <param name="newSourceFullName">新的源文件完整路径。</param>
    /// <returns>是否重新链接成功。</returns>
    bool Relink(string newSourceFullName);

    /// <summary>
    /// 验证链接是否有效。
    /// </summary>
    /// <returns>链接是否有效。</returns>
    bool ValidateLink();

    /// <summary>
    /// 获取链接是否为嵌入对象。
    /// </summary>
    bool IsEmbedded { get; }
}