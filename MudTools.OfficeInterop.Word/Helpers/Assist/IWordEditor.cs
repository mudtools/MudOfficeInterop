//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示有权编辑文档特定部分的单个用户或用户组。
/// <para>注：可授予权限的用户包括单独的参与者以及为"文档工作区"站点定义的用户组。分配给区域和选定内容的权限仅在文档受到保护之后生效。</para>
/// </summary>
public interface IWordEditor : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取用户的显示名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取用户的 ID。
    /// </summary>
    string ID { get; }

    /// <summary>
    /// 获取用户的范围。
    /// </summary>
    IWordRange? Range { get; }

    IWordRange? NextRange { get; }

    /// <summary>
    /// 删除此编辑者权限。
    /// </summary>
    void Delete();

    void DeleteAll();

    void SelectAll();
}