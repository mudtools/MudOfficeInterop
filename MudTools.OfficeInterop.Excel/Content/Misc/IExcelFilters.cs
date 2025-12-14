//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel筛选器接口，表示Excel工作表中的一组筛选条件，支持资源释放和遍历操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelFilters : IDisposable, IEnumerable<IExcelFilter>
{
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取自动筛选器集合中的筛选器数量
    /// </summary>
    int? Count { get; }

    /// <summary>
    /// 根据索引获取自动筛选器（索引从1开始）
    /// </summary>
    /// <param name="index">筛选器索引</param>
    /// <returns>自动筛选器对象</returns>
    IExcelFilter? this[int index] { get; }
}