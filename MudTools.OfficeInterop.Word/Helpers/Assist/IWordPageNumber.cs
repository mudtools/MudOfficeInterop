//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示单个页码。
/// <para>注：PageNumber 对象是 PageNumbers 集合的成员。</para>
/// <para>注：使用 PageNumbers(index)（其中 index 是索引号）可返回单个 PageNumber 对象。</para>
/// </summary>
public interface IWordPageNumber : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中项的索引号。
    /// </summary>
    int Index { get; }

    #endregion

    #region 页码属性 (Page Number Properties)

    /// <summary>
    /// 获取或设置指定页码的对齐方式。
    /// </summary>
    WdPageNumberAlignment Alignment { get; set; }
    #endregion

    #region 页码方法 (Page Number Methods)

    /// <summary>
    /// 删除指定的页码。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将指定的页码复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将指定的页码剪切到剪贴板。
    /// </summary>
    void Cut();

    /// <summary>
    /// 选择指定的页码。
    /// </summary>
    void Select();
    #endregion
}