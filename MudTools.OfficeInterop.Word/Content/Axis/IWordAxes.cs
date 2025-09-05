//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 图表轴集合的封装接口。
/// </summary>
public interface IWordAxes : IEnumerable<IWordAxis>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取轴数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取轴。
    /// </summary>
    /// <param name="type">轴类型。</param>
    /// <param name="axisGroup">轴组。</param>
    IWordAxis this[XlAxisType type, XlAxisGroup axisGroup] { get; }

    /// <summary>
    /// 获取分类轴。
    /// </summary>
    IWordAxis CategoryAxis { get; }

    /// <summary>
    /// 获取数值轴。
    /// </summary>
    IWordAxis ValueAxis { get; }

    /// <summary>
    /// 获取次分类轴。
    /// </summary>
    IWordAxis SecondaryCategoryAxis { get; }

    /// <summary>
    /// 获取次数值轴。
    /// </summary>
    IWordAxis SecondaryValueAxis { get; }

    /// <summary>
    /// 获取所有轴类型列表。
    /// </summary>
    /// <returns>轴类型列表。</returns>
    List<XlAxisType> GetAxisTypes();

    /// <summary>
    /// 获取指定类型的轴数量。
    /// </summary>
    /// <param name="type">轴类型。</param>
    /// <returns>轴数量。</returns>
    int CountByType(XlAxisType type);
}