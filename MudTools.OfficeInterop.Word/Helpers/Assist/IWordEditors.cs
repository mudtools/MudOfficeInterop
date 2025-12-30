//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示有权编辑文档特定部分的用户或用户组的集合。
/// <para>注：使用 <see cref="Add(WdEditorType?)"/> 方法可授予指定用户或组修改文档中的范围或所选内容的权限。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord"), ItemIndex, NoneEnumerable]
public interface IWordEditors : IEnumerable<IWordEditor?>, IOfficeObject<IWordEditors>, IDisposable
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
    /// 获取集合中的编辑者数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引或编辑者ID获取单个编辑者。
    /// </summary>
    /// <param name="index">索引（从1开始）或编辑者的ID。</param>
    /// <returns>指定的编辑者对象。</returns>
    IWordEditor? this[int index] { get; }

    /// <summary>
    /// 通过索引或编辑者ID获取单个编辑者。
    /// </summary>
    /// <param name="id">索引（从1开始）或编辑者的ID。</param>
    /// <returns>指定的编辑者对象。</returns>
    IWordEditor? this[string id] { get; }

    /// <summary>
    /// 添加一个新的编辑者权限。
    /// </summary>
    /// <param name="editorID">要添加的编辑者的ID（可以是用户电子邮件地址或用户组名称）。</param>
    /// <returns>新创建的编辑者对象。</returns>
    IWordEditor? Add(WdEditorType? editorID = null);
}