//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档或节中的一列文本。
/// 封装了 Microsoft.Office.Interop.Word.TextColumn 对象。
/// </summary>
/// <remarks>
/// TextColumn 对象是 TextColumns 集合的成员。每个 TextColumn 对象代表页面或节中的一列文本。
/// </remarks>
public interface IWordTextColumn : IDisposable
{
    #region 属性

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordTextColumn"/> 对象的父对象。
    /// </summary>
    /// <remarks>父对象通常是 TextColumns 集合。</remarks>
    object Parent { get; }

    /// <summary>
    /// 获取或设置文本列的宽度（以磅为单位）。
    /// </summary>
    /// <remarks>
    /// 如果更改某个文本列的宽度，其他文本列的宽度可能会自动调整，以适应页面宽度。
    /// </remarks>
    float? Width { get; set; }

    /// <summary>
    /// 获取或设置文本列右边缘到下一列左边缘的距离（以磅为单位）。
    /// </summary>
    /// <remarks>
    /// 设置此属性会调整文本列之间的间距。
    /// </remarks>
    float? SpaceAfter { get; set; }

    #endregion
}