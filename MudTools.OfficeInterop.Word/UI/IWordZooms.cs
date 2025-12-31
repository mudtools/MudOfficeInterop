//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 中与视图关联的缩放设置集合。
/// 封装了 Microsoft.Office.Interop.Word.Zooms 对象。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord"), ItemIndex, NoneEnumerable]
public interface IWordZooms : IEnumerable<IWordZoom?>, IOfficeObject<IWordZooms>, IDisposable
{

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    /// <remarks>此属性继承自 _IMsoDispObj。</remarks>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建指定对象的应用程序。
    /// </summary>
    /// <remarks>此属性继承自 _IMsoDispObj。</remarks>
    int Creator { get; }

    /// <summary>
    /// 获取代表 <see cref="IWordZooms"/> 对象的父对象。
    /// </summary>
    /// <remarks>对于 Zooms 集合，父对象通常是关联的 View 对象。</remarks>
    object? Parent { get; }

    /// <summary>
    /// 获取指定视图缩放设置对象。
    /// </summary>
    [IgnoreGenerator]
    int Count { get; }

    /// <summary>
    /// 获取指定视图缩放设置对象。
    /// </summary>
    /// <param name="Index"></param>
    /// <returns></returns>
    IWordZoom? this[WdViewType Index] { get; }
}