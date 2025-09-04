//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 ListGallery 对象的集合，这些对象表示“项目符号和编号”对话框中的三个选项卡。
/// <para>注：使用 ListGalleries 属性可返回 ListGalleries 集合。</para>
/// </summary>
public interface IWordListGalleries : IEnumerable<IWordListGallery>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取集合中的 ListGallery 对象数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 <see cref="MsWord.WdListGalleryType"/> 常量获取单个 ListGallery 对象。
    /// </summary>
    /// <param name="index">列表库类型 (<see cref="MsWord.WdListGalleryType"/>)。</param>
    /// <returns>指定的 ListGallery 对象。</returns>
    IWordListGallery this[WdListGalleryType index] { get; }
}