//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示与演示文稿或幻灯片关联的自定义数据集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCustomerData : IEnumerable<IOfficeCustomXMLPart?>, IDisposable
{

    /// <summary>
    /// 获取集合中自定义数据项的数量。
    /// </summary>
    /// <value>集合中自定义数据项的总数。</value>
    int Count { get; }

    /// <summary>
    /// 获取创建此自定义数据集合的应用程序。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取自定义数据集合的父对象。
    /// </summary>
    /// <value>父对象，通常是演示文稿或幻灯片。</value>
    object? Parent { get; }

    /// <summary>
    /// 通过指定的标识符获取自定义XML部件。
    /// </summary>
    /// <param name="id">要获取的自定义XML部件的标识符。</param>
    /// <returns>具有指定标识符的自定义XML部件，如果不存在则为 null。</returns>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    IOfficeCustomXMLPart? this[string id] { get; }

    /// <summary>
    /// 向集合添加一个新的自定义XML部件。
    /// </summary>
    /// <returns>新添加的自定义XML部件对象。</returns>
    IOfficeCustomXMLPart? Add();

    /// <summary>
    /// 从集合中删除具有指定标识符的自定义XML部件。
    /// </summary>
    /// <param name="id">要删除的自定义XML部件的标识符。</param>
    void Delete(string id);
}