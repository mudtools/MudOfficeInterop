//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel调整值集合的接口，提供对一系列浮点调整值的访问和管理功能
/// </summary>
/// <remarks>
/// 该接口继承自IEnumerable&lt;float&gt;和IDisposable，支持迭代访问和资源释放
/// </remarks>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelAdjustments : IEnumerable<float>, IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取集合中元素的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置指定索引处的调整值
    /// </summary>
    /// <param name="index">要获取或设置的元素从零开始的索引</param>
    /// <returns>指定索引处的浮点调整值</returns>
    float this[int index] { get; set; }
}