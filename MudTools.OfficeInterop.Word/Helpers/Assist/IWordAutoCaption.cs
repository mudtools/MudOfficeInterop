//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示一个自动图文集条目。
/// <para>注：AutoCaption 对象是 AutoCaptions 集合的成员。</para>
/// <para>注：使用 AutoCaptions(index) 可返回单个 AutoCaption 对象，其中 index 是类名或索引号。</para>
/// </summary>
public interface IWordAutoCaption : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中项的索引号。
    /// </summary>
    int Index { get; }

    #endregion

    #region 自动题注属性 (AutoCaption Properties)

    /// <summary>
    /// 获取或设置一个值，该值指示是否自动为指定项目添加题注。
    /// </summary>
    bool AutoInsert { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示在自动插入的题注中使用的题注标签。
    /// </summary>
    string CaptionLabel { get; set; }

    /// <summary>
    /// 获取表示指定项目类名的字符串。
    /// </summary>
    string Name { get; }

    #endregion
}