//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Borders 的接口，用于操作边框集合。
/// </summary>
public interface IWordBorders : IEnumerable<IWordBorder>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取边框的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据边框类型获取边框。
    /// </summary>
    IWordBorder this[WdBorderType borderType] { get; }

    /// <summary>
    /// 获取或设置是否启用边框。
    /// </summary>
    bool Enable { get; set; }

    /// <summary>
    /// 应用边框样式到所有边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条粗细。</param>
    /// <param name="color">颜色。</param>
    void ApplyStyle(WdLineStyle lineStyle, WdLineWidth lineWidth, WdColor color);

    /// <summary>
    /// 获取指定类型的边框是否存在。
    /// </summary>
    /// <param name="borderType">边框类型。</param>
    /// <returns>是否存在。</returns>
    bool Contains(WdBorderType borderType);

    /// <summary>
    /// 获取所有边框类型的列表。
    /// </summary>
    /// <returns>边框类型列表。</returns>
    List<WdBorderType> GetBorderTypes();
}