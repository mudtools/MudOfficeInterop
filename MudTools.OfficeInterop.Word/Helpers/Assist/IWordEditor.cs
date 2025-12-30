//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示已被授予特定权限以编辑文档部分的单个用户。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordEditor : IOfficeObject<IWordEditor>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
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

    /// <summary>
    /// 获取用户的下一个范围。
    /// </summary>
    IWordRange? NextRange { get; }

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 删除文档中特定用户的所有编辑权限。
    /// </summary>
    void DeleteAll();

    /// <summary>
    /// 选择主故事、画布或文档页眉页脚中的所有形状。
    /// </summary>
    void SelectAll();
}