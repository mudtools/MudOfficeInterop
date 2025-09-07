//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示单个页眉或页脚。
/// <para>注：HeaderFooter 对象是 HeadersFooters 集合的成员。</para>
/// <para>注：使用 HeadersFooters(index)（其中 index 是页眉或页脚的索引号）可返回单个 HeaderFooter 对象。</para>
/// <para>注：索引号 1 代表首页页眉或页脚，索引号 2 代表奇数页页眉或页脚，索引号 3 代表偶数页页眉或页脚。</para>
/// </summary>
public interface IWordHeaderFooter : IDisposable
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

    #endregion

    #region 页眉/页脚属性 (Header/Footer Properties)

    /// <summary>
    /// 获取表示页眉或页脚中文本的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在页眉或页脚中显示页码。
    /// </summary>
    IWordPageNumbers? PageNumbers { get; }

    /// <summary>
    /// 获取表示页眉或页脚类型的 WdHeaderFooterIndex 常量。
    /// </summary>
    WdHeaderFooterIndex Index { get; }

    /// <summary>
    /// 获取表示页眉或页脚所在节的 Shapes 集合。
    /// </summary>
    IWordShapes? Shapes { get; }
    #endregion
}